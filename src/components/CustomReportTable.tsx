import React, { useState, useMemo, useRef, useEffect } from 'react';
import { 
  FileSpreadsheet, Printer, RotateCcw, 
  Search, CheckSquare, Square, ArrowUpDown, ArrowUp, ArrowDown,
  Edit3, Check, Shield, Columns, Minimize2, Maximize2,
  Building2, FileDown, Loader2
} from 'lucide-react';
import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { ImmRecord, MovementData, MasterItem, CRTColumnId } from '../types';
import { CRT_COLUMNS_DEFINITION, ALL_COLUMN_IDS } from '../utils/crtPresets';
import { CustomReportColumnModal } from './CustomReportColumnModal';
import { logActivity } from '../utils/activityLogger';
import { formatToDDMMYYYY } from '../App';

interface CustomReportTableProps {
  records: ImmRecord[];
  movementMap: Record<string, MovementData>;
  currentUser?: { name: string; title: string } | null;
  masterData?: MasterItem[];
  showToast: (msg: string) => void;
}

export interface CustomReportRow {
  id: string;
  passport: string;
  nationality: string;
  fullname: string;
  gender: 'M' | 'F' | '';
  visaType: string;
  visaNumber?: string;
  dob?: string;
  stayFrom: string;
  stayTo: string;
  lastArrivalDate: string;
  arrivalTimestampMs: number;
  elapsedDays: number;
  address: string;
  remarks: string;
  selected: boolean;
  orderIndex: number;
}

export const CustomReportTable: React.FC<CustomReportTableProps> = ({
  records,
  movementMap,
  currentUser,
  masterData = [],
  showToast
}) => {
  const today = new Date().toLocaleDateString('en-CA'); // 'YYYY-MM-DD' in local time

  // Table Title & Date Configuration
  const [tableTitle, setTableTitle] = useState<string>(() => 
    localStorage.getItem('crt_tableTitle') || 'လက်ရှိ နေထိုင်ဆဲ နိုင်ငံခြားသားများ စာရင်း'
  );
  const [isEditingTitle, setIsEditingTitle] = useState(false);
  const [selectedDate, setSelectedDate] = useState<string>(today);
  const [useBurmeseDigits, setUseBurmeseDigits] = useState<boolean>(true);

  // Officer Signature Configuration
  const [showSigner, setShowSigner] = useState<boolean>(() => {
    const saved = localStorage.getItem('crt_showSigner');
    return saved !== null ? saved === 'true' : true;
  });
  const [officerTitle, setOfficerTitle] = useState<string>(() => 
    localStorage.getItem('crt_officerTitle') || currentUser?.title || 'လဝကမှူး'
  );
  const [officerName, setOfficerName] = useState<string>(() => 
    localStorage.getItem('crt_officerName') || currentUser?.name || 'ဒုတိယလဝကမှူး'
  );

  const handleToggleShowSigner = (val: boolean) => {
    setShowSigner(val);
    try {
      localStorage.setItem('crt_showSigner', String(val));
    } catch {}
    showToast(val ? "Signer လက်မှတ် ထည့်သွင်းထားပါသည်" : "Signer လက်မှတ် ဖြုတ်ထားပါသည်");
  };

  const handleSaveTitle = (newTitle: string) => {
    const trimmed = newTitle.trim() || 'Custom Report Table';
    setTableTitle(trimmed);
    setIsEditingTitle(false);
    try {
      localStorage.setItem('crt_tableTitle', trimmed);
    } catch {}
    showToast("ခေါင်းစဉ် အသစ် ပြောင်းလဲပြီးပါပြီ");
  };

  const handleSaveOfficerTitle = (val: string) => {
    setOfficerTitle(val);
    try {
      localStorage.setItem('crt_officerTitle', val);
    } catch {}
  };

  const handleSaveOfficerName = (val: string) => {
    setOfficerName(val);
    try {
      localStorage.setItem('crt_officerName', val);
    } catch {}
  };

  // Direct On-Screen Filter States
  const [dataPool, setDataPool] = useState<'STILL_IN' | 'ALL_MOVEMENTS' | 'INBOUND' | 'OUTBOUND'>('STILL_IN');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [filterNationality, setFilterNationality] = useState<string>('ALL');
  const [filterVisaType, setFilterVisaType] = useState<string>('ALL');
  const [filterAddress, setFilterAddress] = useState<string>('ALL');
  const [filterGender, setFilterGender] = useState<'ALL' | 'M' | 'F'>('ALL');
  const [filterAddressType, setFilterAddressType] = useState<'ALL' | 'COMPANY' | 'OTHER'>('ALL');
  const [excludeVisaS, setExcludeVisaS] = useState<boolean>(false);
  const [minElapsedDays, setMinElapsedDays] = useState<string>('');
  const [startDate, setStartDate] = useState<string>('');
  const [endDate, setEndDate] = useState<string>('');

  // Preset: Daily Hotel List (No Visa S) - နေ့စဥ်ပို့ရန် ဟိုတယ်စာရင်း
  const applyDailyHotelPreset = () => {
    setTableTitle('လက်ရှိ ဟိုတယ်နေနိုင်ငံခြားသားစာရင်း');
    setDataPool('STILL_IN');
    setFilterAddressType('OTHER');
    setExcludeVisaS(true);
    setFilterNationality('ALL');
    setFilterVisaType('ALL');
    setFilterAddress('ALL');
    setFilterGender('ALL');
    setMinElapsedDays('');
    setStartDate('');
    setEndDate('');
    setSearchQuery('');
    setSortBy('elapsed');
    setSortDirection('desc');
    try {
      localStorage.setItem('crt_tableTitle', 'လက်ရှိ ဟိုတယ်နေနိုင်ငံခြားသားစာရင်း');
    } catch {}
    showToast("နေ့စဥ်ပို့ရန် ဟိုတယ်စာရင်း (ဗီဇာ S မပါ) ရွေးချယ်သတ်မှတ်ပြီးပါပြီ");
  };

  // Sorting States
  const [sortBy, setSortBy] = useState<
    'elapsed' | 'arrival' | 'passport' | 'name' | 'nationality' | 'address' | 'visa' | 'stayTo' | 'custom'
  >('elapsed');
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('desc');

  // Manual Row Selection & Order
  const [selectedPassports, setSelectedPassports] = useState<Record<string, boolean>>({});
  const [manualRowOrder, setManualRowOrder] = useState<string[]>([]);
  const [customRemarks, setCustomRemarks] = useState<Record<string, string>>({});
  const [isExportingPDF, setIsExportingPDF] = useState<boolean>(false);

  // Column Visibility
  const [visibleColumns, setVisibleColumns] = useState<CRTColumnId[]>(() => {
    try {
      const raw = localStorage.getItem('crt_visible_columns_v2');
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch {}
    return [...ALL_COLUMN_IDS];
  });
  const [isColumnModalOpen, setIsColumnModalOpen] = useState<boolean>(false);

  const handleToggleColumn = (colId: CRTColumnId) => {
    setVisibleColumns(prev => {
      let updated: CRTColumnId[];
      if (prev.includes(colId)) {
        if (prev.length <= 1) {
          showToast("အနည်းဆုံး ကော်လံ (၁) ခု ဖွင့်ထားရပါမည်");
          return prev;
        }
        updated = prev.filter(c => c !== colId);
      } else {
        updated = ALL_COLUMN_IDS.filter(c => prev.includes(c) || c === colId);
      }
      try {
        localStorage.setItem('crt_visible_columns_v2', JSON.stringify(updated));
      } catch {}
      return updated;
    });
  };

  const handleSelectAllColumns = () => {
    setVisibleColumns([...ALL_COLUMN_IDS]);
    try {
      localStorage.setItem('crt_visible_columns_v2', JSON.stringify(ALL_COLUMN_IDS));
    } catch {}
    showToast("ကော်လံအားလုံး ဖွင့်လှစ်ပြီးပါပြီ");
  };

  const handleResetDefaultColumns = () => {
    setVisibleColumns([...ALL_COLUMN_IDS]);
    try {
      localStorage.setItem('crt_visible_columns_v2', JSON.stringify(ALL_COLUMN_IDS));
    } catch {}
    showToast("မူလကော်လံများ ပြန်လည်သတ်မှတ်ပြီးပါပြီ");
  };

  // Zoom / Scale State for A4 Landscape preview
  const [viewZoomMode, setViewZoomMode] = useState<'fit' | 'full'>('fit');
  const [fitScale, setFitScale] = useState<number>(1);
  const previewContainerRef = useRef<HTMLDivElement>(null);
  const printRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const calcFitScale = () => {
      if (!previewContainerRef.current) return;
      const containerWidth = previewContainerRef.current.clientWidth - 32;
      const paperWidthPx = 1122; // 297mm @ 96dpi approx
      if (containerWidth > 0 && containerWidth < paperWidthPx) {
        const scale = Math.max(0.35, Math.min(1, containerWidth / paperWidthPx));
        setFitScale(scale);
      } else {
        setFitScale(1);
      }
    };
    calcFitScale();
    window.addEventListener('resize', calcFitScale);
    return () => window.removeEventListener('resize', calcFitScale);
  }, []);

  const currentScale = viewZoomMode === 'fit' ? fitScale : 1;

  // Number Conversion Helpers
  const toBurmeseDigits = (n: number | string): string => {
    const burmeseNums = ['၀', '၁', '၂', '၃', '၄', '၅', '၆', '၇', '၈', '၉'];
    return String(n).replace(/[0-9]/g, d => burmeseNums[parseInt(d, 10)] || d);
  };

  const toBurmeseSlashDate = (isoDate: string | undefined): string => {
    if (!isoDate) return '-';
    const formatted = formatToDDMMYYYY(isoDate);
    return toBurmeseDigits(formatted);
  };

  // Parse any date/timestamp into milliseconds safely
  const parseDateToMs = (dateStr?: string | number): number => {
    if (!dateStr) return 0;
    if (typeof dateStr === 'number') return dateStr;
    const str = String(dateStr).trim();
    if (/^\d+$/.test(str)) return parseInt(str, 10);
    const parsed = Date.parse(str);
    if (!isNaN(parsed)) return parsed;
    // Format DD-MM-YYYY or DD/MM/YYYY
    const parts = str.split(/[-\/]/);
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        return new Date(`${parts[0]}-${parts[1]}-${parts[2]}`).getTime() || 0;
      }
      return new Date(`${parts[2]}-${parts[1]}-${parts[0]}`).getTime() || 0;
    }
    return 0;
  };

  // Base Data Extraction matching exact Still In Movement Log logic & reference date
  const basePoolRows: CustomReportRow[] = useMemo(() => {
    const selectedDateStr = selectedDate || today;
    const isCurrentRealtime = !selectedDate || selectedDate === today;

    if (dataPool === 'STILL_IN') {
      const activePassports = Object.keys(movementMap || {}).filter(pp => {
        const mov = movementMap[pp];
        if (!mov) return false;

        // Current real-time still-in (today):
        // Exactly matches Movement Log logic: !mov.out
        if (isCurrentRealtime) {
          return !mov.out;
        }

        // Historical / Specific reference date:
        // 1. Must have arrival on or before selectedDate
        const arrDateStr = mov.in ? (mov.in.includes('T') ? mov.in.split('T')[0] : mov.in.split(' ')[0]) : '';
        if (arrDateStr && arrDateStr > selectedDateStr) {
          return false; // Arrived after reference date
        }

        // 2. Must NOT have departed on or before selectedDate
        const depDateStr = mov.out ? (mov.out.includes('T') ? mov.out.split('T')[0] : mov.out.split(' ')[0]) : '';
        if (depDateStr && depDateStr <= selectedDateStr) {
          return false; // Departed on or before reference date
        }

        return true;
      });

      return activePassports.map((pp, idx) => {
        const mov = movementMap[pp];
        const arrDate = mov.in ? (mov.in.includes('T') ? mov.in.split('T')[0] : mov.in.split(' ')[0]) : '';
        const inTime = mov.inTime || parseDateToMs(mov.in);
        
        let elapsed = 0;
        if (inTime > 0) {
          const refTime = new Date(selectedDateStr).getTime();
          const diffMs = refTime - inTime;
          elapsed = Math.max(0, Math.floor(diffMs / (1000 * 60 * 60 * 24)));
        } else if (arrDate) {
          const aMs = parseDateToMs(arrDate);
          const rMs = parseDateToMs(selectedDateStr);
          if (aMs && rMs) {
            elapsed = Math.max(0, Math.floor((rMs - aMs) / (1000 * 60 * 60 * 24)));
          }
        }

        const addr = mov.loc || '-';
        const nat = mov.nat || '-';
        const vType = mov.visa || '-';
        const name = mov.n || '-';
        const gen = (mov.gender === 'M' || mov.gender === 'F') ? mov.gender : '';

        return {
          id: pp,
          passport: pp,
          nationality: nat,
          fullname: name,
          gender: gen,
          visaType: vType,
          visaNumber: mov.visaNumber || '',
          dob: mov.dob || '',
          stayFrom: mov.start || '',
          stayTo: mov.end || '',
          lastArrivalDate: arrDate ? (useBurmeseDigits ? toBurmeseSlashDate(arrDate) : formatToDDMMYYYY(arrDate)) : '-',
          arrivalTimestampMs: inTime,
          elapsedDays: elapsed,
          address: addr,
          remarks: customRemarks[pp] || '',
          selected: selectedPassports[pp] !== false,
          orderIndex: idx
        };
      });
    }

    // ALL_MOVEMENTS / INBOUND / OUTBOUND
    let filteredRecords = [...(records || [])];
    if (dataPool === 'INBOUND') {
      filteredRecords = filteredRecords.filter(r => r.direction === 'IN' || r.type === 'Arrival' || r.mode === 'IN' || !r.direction);
    } else if (dataPool === 'OUTBOUND') {
      filteredRecords = filteredRecords.filter(r => r.direction === 'OUT' || r.type === 'Departure' || r.mode === 'OUT');
    }

    return filteredRecords.map((r, idx) => {
      const pp = r.passportNo || r.passport || `REC-${idx}`;
      const arrDate = r.flightDate || r.formC?.date || r.date || (r.timestamp ? r.timestamp.split('T')[0] : '');
      const inTime = r.timestamp ? parseDateToMs(r.timestamp) : 0;

      let elapsed = 0;
      if (inTime > 0) {
        const refTime = new Date(selectedDateStr).getTime();
        const diffMs = refTime - inTime;
        elapsed = Math.max(0, Math.floor(diffMs / (1000 * 60 * 60 * 24)));
      }

      const addr = r.formC?.address || r.address || '-';
      const nat = r.nationality || '-';
      const vType = r.visaType || '-';
      const name = r.fullName || r.name || '-';
      const gen = (r.gender === 'M' || r.gender === 'F') ? r.gender : '';

      return {
        id: `${pp}-${idx}`,
        passport: pp,
        nationality: nat,
        fullname: name,
        gender: gen,
        visaType: vType,
        visaNumber: r.visaNumber || r.formC?.visaNumber || '',
        dob: r.dob || '',
        stayFrom: r.stayPermitFrom || r.formC?.stayPermitFrom || '',
        stayTo: r.stayPermitTo || r.formC?.stayPermitTo || '',
        lastArrivalDate: arrDate ? (useBurmeseDigits ? toBurmeseSlashDate(arrDate) : formatToDDMMYYYY(arrDate)) : '-',
        arrivalTimestampMs: inTime,
        elapsedDays: elapsed,
        address: addr,
        remarks: customRemarks[pp] || '',
        selected: selectedPassports[pp] !== false,
        orderIndex: idx
      };
    });
  }, [records, movementMap, dataPool, selectedDate, today, useBurmeseDigits, selectedPassports, customRemarks]);

  // Unique Filter Options
  const uniqueNationalities = useMemo(() => {
    const set = new Set<string>();
    basePoolRows.forEach(r => { if (r.nationality && r.nationality !== '-') set.add(r.nationality); });
    return Array.from(set).sort();
  }, [basePoolRows]);

  const uniqueVisaTypes = useMemo(() => {
    const set = new Set<string>();
    basePoolRows.forEach(r => { if (r.visaType && r.visaType !== '-') set.add(r.visaType); });
    return Array.from(set).sort();
  }, [basePoolRows]);

  const uniqueAddresses = useMemo(() => {
    const set = new Set<string>();
    basePoolRows.forEach(r => { if (r.address && r.address !== '-') set.add(r.address); });
    return Array.from(set).sort();
  }, [basePoolRows]);

  // Filtered & Sorted Rows
  const sortedRows = useMemo(() => {
    let rows = basePoolRows.filter(r => {
      // Live search
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const match = 
          r.passport.toLowerCase().includes(q) ||
          r.fullname.toLowerCase().includes(q) ||
          r.nationality.toLowerCase().includes(q) ||
          r.address.toLowerCase().includes(q) ||
          r.visaType.toLowerCase().includes(q);
        if (!match) return false;
      }

      // Nationality
      if (filterNationality !== 'ALL' && r.nationality !== filterNationality) return false;

      // Visa Type
      if (filterVisaType !== 'ALL' && r.visaType !== filterVisaType) return false;

      // Address
      if (filterAddress !== 'ALL' && r.address !== filterAddress) return false;

      // Address category (Company vs Other/Hotel)
      if (filterAddressType === 'COMPANY') {
        const isCo = /co\.|ltd|company|ကုမ္ပဏီ|မိုင်း|စက်ရုံ/i.test(r.address);
        if (!isCo) return false;
      } else if (filterAddressType === 'OTHER') {
        const isCo = /co\.|ltd|company|ကုမ္ပဏီ|မိုင်း|စက်ရုံ/i.test(r.address);
        if (isCo) return false;
      }

      // Exclude Visa S (Special / Stay / S-)
      if (excludeVisaS) {
        const vClean = (r.visaType || '').trim();
        // Check if visa type starts with 'S' (case-insensitive) or has 'special' or 'stay'
        if (/^(s|special|stay)/i.test(vClean)) {
          return false;
        }
      }

      // Gender
      if (filterGender !== 'ALL' && r.gender !== filterGender) return false;

      // Min elapsed days
      if (minElapsedDays.trim() !== '') {
        const minDays = parseInt(minElapsedDays, 10);
        if (!isNaN(minDays) && r.elapsedDays < minDays) return false;
      }

      // Date Range
      if (startDate || endDate) {
        const dateStr = r.lastArrivalDate;
        if (dateStr && dateStr !== '-') {
          const parts = dateStr.split('/');
          if (parts.length === 3) {
            const iso = `${parts[2]}-${parts[1]}-${parts[0]}`;
            if (startDate && iso < startDate) return false;
            if (endDate && iso > endDate) return false;
          }
        }
      }

      return true;
    });

    // Custom Order handling
    const customOrderMap = new Map<string, number>();
    manualRowOrder.forEach((id, idx) => customOrderMap.set(id, idx));

    rows.sort((a, b) => {
      if (sortBy === 'custom') {
        const orderA = customOrderMap.has(a.id) ? customOrderMap.get(a.id)! : a.orderIndex;
        const orderB = customOrderMap.has(b.id) ? customOrderMap.get(b.id)! : b.orderIndex;
        return sortDirection === 'asc' ? orderA - orderB : orderB - orderA;
      }
      if (sortBy === 'elapsed') {
        return sortDirection === 'asc' ? a.elapsedDays - b.elapsedDays : b.elapsedDays - a.elapsedDays;
      }
      if (sortBy === 'arrival') {
        return sortDirection === 'asc' ? a.arrivalTimestampMs - b.arrivalTimestampMs : b.arrivalTimestampMs - a.arrivalTimestampMs;
      }
      if (sortBy === 'passport') {
        return sortDirection === 'asc' ? a.passport.localeCompare(b.passport) : b.passport.localeCompare(a.passport);
      }
      if (sortBy === 'name') {
        return sortDirection === 'asc' ? a.fullname.localeCompare(b.fullname) : b.fullname.localeCompare(a.fullname);
      }
      if (sortBy === 'nationality') {
        return sortDirection === 'asc' ? a.nationality.localeCompare(b.nationality) : b.nationality.localeCompare(a.nationality);
      }
      if (sortBy === 'address') {
        return sortDirection === 'asc' ? a.address.localeCompare(b.address) : b.address.localeCompare(a.address);
      }
      if (sortBy === 'visa') {
        return sortDirection === 'asc' ? a.visaType.localeCompare(b.visaType) : b.visaType.localeCompare(a.visaType);
      }
      if (sortBy === 'stayTo') {
        return sortDirection === 'asc' ? (a.stayTo || '').localeCompare(b.stayTo || '') : (b.stayTo || '').localeCompare(a.stayTo || '');
      }
      return 0;
    });

    return rows;
  }, [
    basePoolRows, searchQuery, filterNationality, filterVisaType, filterAddress,
    filterAddressType, excludeVisaS, filterGender, minElapsedDays, startDate, endDate,
    sortBy, sortDirection, manualRowOrder
  ]);

  // Display only selected rows
  const displayRows = useMemo(() => {
    return sortedRows.filter(r => r.selected);
  }, [sortedRows]);

  // Paginate displayRows into 12 rows per A4 landscape page chunk
  const ROWS_PER_PAGE = 12;
  const paginatedPages = useMemo(() => {
    if (displayRows.length === 0) return [[]];
    const pages: CustomReportRow[][] = [];
    for (let i = 0; i < displayRows.length; i += ROWS_PER_PAGE) {
      pages.push(displayRows.slice(i, i + ROWS_PER_PAGE));
    }
    return pages;
  }, [displayRows]);

  // Toggle selection
  const toggleRowSelect = (id: string) => {
    setSelectedPassports(prev => ({
      ...prev,
      [id]: prev[id] === false ? true : false
    }));
  };

  const toggleSelectAll = (select: boolean) => {
    const updated: Record<string, boolean> = {};
    sortedRows.forEach(r => { updated[r.id] = select; });
    setSelectedPassports(prev => ({ ...prev, ...updated }));
    showToast(select ? "စာရင်းအားလုံး ရွေးချယ်ပြီးပါပြီ" : "ရွေးချယ်မှု အားလုံး ပယ်ဖျက်ပြီးပါပြီ");
  };

  // Reordering rows
  const handleMoveRow = (id: string, direction: 'up' | 'down') => {
    const currentOrder = manualRowOrder.length > 0 
      ? [...manualRowOrder] 
      : sortedRows.map(r => r.id);

    const idx = currentOrder.indexOf(id);
    if (idx === -1) return;

    if (direction === 'up' && idx > 0) {
      const temp = currentOrder[idx];
      currentOrder[idx] = currentOrder[idx - 1];
      currentOrder[idx - 1] = temp;
      setManualRowOrder(currentOrder);
      setSortBy('custom');
    } else if (direction === 'down' && idx < currentOrder.length - 1) {
      const temp = currentOrder[idx];
      currentOrder[idx] = currentOrder[idx + 1];
      currentOrder[idx + 1] = temp;
      setManualRowOrder(currentOrder);
      setSortBy('custom');
    }
  };

  // Header click sorting
  const handleHeaderSort = (field: typeof sortBy) => {
    if (sortBy === field) {
      setSortDirection(prev => prev === 'asc' ? 'desc' : 'asc');
    } else {
      setSortBy(field);
      setSortDirection('asc');
    }
  };

  // Excel Export
  const exportToExcel = () => {
    if (displayRows.length === 0) {
      showToast("Export ပြုလုပ်ရန် စာရင်းမရှိပါ");
      return;
    }

    const headers: string[] = [];
    if (visibleColumns.includes('sr')) headers.push('စဉ်');
    if (visibleColumns.includes('passport')) headers.push('နိုင်ငံကူးလက်မှတ်အမှတ်');
    if (visibleColumns.includes('nationality')) headers.push('နိုင်ငံအမည်');
    if (visibleColumns.includes('fullname')) headers.push('အမည်');
    if (visibleColumns.includes('gender')) headers.push('ကျား/မ');
    if (visibleColumns.includes('visaType')) headers.push('ဗီဇာအမျိုးအစား');
    if (visibleColumns.includes('stayPeriod')) headers.push('ဗီဇာသက်တမ်း (မှ/ထိ)');
    if (visibleColumns.includes('arrivalDate')) headers.push('နောက်ဆုံးရောက်ရှိရက်စွဲ');
    if (visibleColumns.includes('elapsedDays')) headers.push('နောက်ဆုံးရောက်ရှိသည့်နေ့မှ ရက်ပေါင်း');
    if (visibleColumns.includes('address')) headers.push('တည်းခိုလိပ်စာ');
    if (visibleColumns.includes('remarks')) headers.push('မှတ်ချက်');

    const dataMatrix = displayRows.map((r, idx) => {
      const row: (string | number)[] = [];
      if (visibleColumns.includes('sr')) row.push(idx + 1);
      if (visibleColumns.includes('passport')) row.push(r.passport);
      if (visibleColumns.includes('nationality')) row.push(r.nationality);
      if (visibleColumns.includes('fullname')) row.push(r.fullname);
      if (visibleColumns.includes('gender')) row.push(r.gender === 'M' ? 'ကျား' : r.gender === 'F' ? 'မ' : '-');
      if (visibleColumns.includes('visaType')) row.push(r.visaType);
      if (visibleColumns.includes('stayPeriod')) row.push((r.stayFrom || r.stayTo) ? `${r.stayFrom || '-'} မှ ${r.stayTo || '-'}` : '-');
      if (visibleColumns.includes('arrivalDate')) row.push(r.lastArrivalDate);
      if (visibleColumns.includes('elapsedDays')) row.push(r.elapsedDays);
      if (visibleColumns.includes('address')) row.push(r.address);
      if (visibleColumns.includes('remarks')) row.push(r.remarks || '');
      return row;
    });

    const ws = XLSX.utils.aoa_to_sheet([
      [tableTitle],
      [`ရက်စွဲ: ${selectedDate}`],
      [],
      headers,
      ...dataMatrix,
      [],
      ...(showSigner ? [[], ['', '', '', '', '', '', '', '', officerTitle], ['', '', '', '', '', '', '', '', officerName]] : [])
    ]);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Custom_Report");
    const cleanFileName = `${(tableTitle || 'Custom_Report').replace(/\s+/g, '_')}_${selectedDate}.xlsx`;
    XLSX.writeFile(wb, cleanFileName);

    logActivity({
      action: 'SETTINGS_CHANGE',
      module: 'CRT',
      targetId: 'Custom_Report',
      details: `Exported Custom Report Excel (${displayRows.length} rows)`
    });

    showToast("EXCEL ဖိုင် ဒေါင်းလုဒ်ရယူပြီးပါပြီ");
  };

  // Direct High-Fidelity A4 Landscape PDF Export (Pyidaungsu Font 100% Compatible)
  const exportDirectPDF = async () => {
    if (!printRef.current) return;
    if (displayRows.length === 0) {
      showToast("PDF ထုတ်ယူရန် စာရင်းမရှိပါ");
      return;
    }

    setIsExportingPDF(true);
    showToast("A4 Landscape PDF ပြင်ဆင်နေပါသည်...");

    // Store original styles to restore in finally block
    const styleTags = Array.from(document.querySelectorAll('style'));
    const originalStyles = styleTags.map(tag => ({ tag, html: tag.innerHTML }));

    // Temporarily replace oklch(...) colors inside styles to prevent html2canvas crashes
    styleTags.forEach(tag => {
      if (tag.innerHTML.includes('oklch')) {
        tag.innerHTML = tag.innerHTML.replace(/oklch\([^)]+\)/g, 'rgb(75, 85, 99)');
      }
    });

    try {
      await new Promise(resolve => setTimeout(resolve, 250));

      const pageElements = printRef.current.querySelectorAll('.crt-pdf-page');
      if (!pageElements || pageElements.length === 0) {
        showToast("⚠️ PDF စာမျက်နှာ ရှာမတွေ့ပါ");
        setIsExportingPDF(false);
        return;
      }

      const pdf = new jsPDF({
        orientation: 'landscape',
        unit: 'mm',
        format: 'a4',
        compress: true
      });

      for (let i = 0; i < pageElements.length; i++) {
        const pageEl = pageElements[i] as HTMLElement;

        // Render each exact A4 page element with high resolution
        const canvas = await html2canvas(pageEl, {
          scale: 2.2, // Crisp 300dpi-equivalent resolution
          useCORS: true,
          logging: false,
          backgroundColor: '#ffffff',
          windowWidth: 1200
        });

        const imgData = canvas.toDataURL('image/jpeg', 0.98);

        if (i > 0) {
          pdf.addPage('a4', 'landscape');
        }

        // Exact A4 landscape placement: 297mm x 210mm
        pdf.addImage(imgData, 'JPEG', 0, 0, 297, 210, undefined, 'FAST');
      }

      const safeTitle = (tableTitle || 'Custom_Report').replace(/[^a-zA-Z0-9_\u1000-\u109F]/g, '_');
      pdf.save(`${safeTitle}_${selectedDate}.pdf`);

      logActivity({
        action: 'SETTINGS_CHANGE',
        module: 'CRT',
        targetId: 'Custom_Report_PDF',
        details: `Downloaded Direct A4 Landscape PDF (${displayRows.length} rows, ${pageElements.length} pages)`
      });

      showToast("A4 LANDSCAPE PDF ဒေါင်းလုဒ် ရယူပြီးပါပြီ");
    } catch (e) {
      console.error("PDF Export Error:", e);
      showToast("⚠️ PDF ထုတ်ယူရာတွင် ချို့ယွင်းချက်ရှိပါသည်");
    } finally {
      originalStyles.forEach(({ tag, html }) => {
        tag.innerHTML = html;
      });
      setIsExportingPDF(false);
    }
  };

  // Browser Print
  const triggerPrint = () => {
    window.print();
  };

  return (
    <div className="max-w-7xl mx-auto py-6 px-4 space-y-6">
      {/* Sleek Minimal Header Banner */}
      <div className="bg-slate-900 text-white rounded-3xl p-6 sm:p-7 shadow-xl border border-slate-800 no-print space-y-5">
        <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4">
          {/* Editable Title */}
          <div className="flex items-center gap-3 flex-1 min-w-0">
            <div className="p-2.5 bg-indigo-600/30 text-indigo-400 rounded-2xl border border-indigo-500/30 shrink-0">
              <FileSpreadsheet size={24} />
            </div>
            <div className="flex-1 min-w-0">
              {isEditingTitle ? (
                <div className="flex items-center gap-2 max-w-xl">
                  <input
                    type="text"
                    defaultValue={tableTitle}
                    id="crt_title_input"
                    className="flex-1 bg-slate-800 text-white border border-indigo-400 px-3 py-1.5 rounded-xl font-bold text-base outline-none font-pyidaungsu"
                    autoFocus
                  />
                  <button
                    onClick={() => {
                      const input = document.getElementById('crt_title_input') as HTMLInputElement;
                      if (input) handleSaveTitle(input.value);
                    }}
                    className="p-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl cursor-pointer shrink-0 shadow-md"
                    title="သိမ်းဆည်းမည်"
                  >
                    <Check size={16} />
                  </button>
                </div>
              ) : (
                <div 
                  onClick={() => setIsEditingTitle(true)}
                  className="group cursor-pointer flex items-center gap-2"
                  title="ခေါင်းစဉ်ပြင်ရန် နှိပ်ပါ"
                >
                  <h2 className="text-xl sm:text-2xl font-black uppercase tracking-tight font-pyidaungsu text-white group-hover:text-indigo-200 transition-colors truncate">
                    {tableTitle}
                  </h2>
                  <Edit3 size={15} className="text-indigo-400 opacity-60 group-hover:opacity-100 shrink-0" />
                </div>
              )}
              <p className="text-xs text-slate-400 font-semibold mt-0.5">
                Official A4 Landscape Custom Report System
              </p>
            </div>
          </div>

          {/* Quick Actions & Output Controls */}
          <div className="flex flex-wrap items-center gap-2.5">
            {/* Dedicated Preset: နေ့စဥ်ပို့ရန် ဟိုတယ်စာရင်း */}
            <button
              onClick={applyDailyHotelPreset}
              className="px-3.5 py-2 text-xs font-black rounded-xl bg-amber-500 hover:bg-amber-400 text-slate-950 shadow-md transition-all flex items-center gap-1.5 cursor-pointer active:scale-95"
              title="နေ့စဥ်ပို့ရန် ဟိုတယ်စာရင်း (ဗီဇာ S မပါ) အလိုအလျောက် သတ်မှတ်မည်"
            >
              <Building2 size={14} />
              <span>နေ့စဥ်ပို့ရန် ဟိုတယ်စာရင်း</span>
            </button>

            {/* Reference Date Input */}
            <div className="flex items-center gap-1.5 bg-slate-800/80 border border-slate-700 px-3 py-1.5 rounded-xl">
              <span className="text-[10px] text-slate-400 font-bold uppercase">ရက်စွဲ:</span>
              <input
                type="date"
                value={selectedDate}
                onChange={(e) => setSelectedDate(e.target.value)}
                className="bg-transparent text-white font-bold text-xs outline-none cursor-pointer"
              />
            </div>

            {/* Burmese Digits Toggle */}
            <button
              onClick={() => setUseBurmeseDigits(!useBurmeseDigits)}
              className={`px-3 py-1.5 text-xs font-black rounded-xl border transition-all cursor-pointer ${
                useBurmeseDigits 
                  ? 'bg-amber-500/20 text-amber-300 border-amber-500/40' 
                  : 'bg-slate-800 text-slate-300 border-slate-700'
              }`}
            >
              <span>{useBurmeseDigits ? '၁,၂,၃ (မြန်မာ)' : '1,2,3 (အင်္ဂလိပ်)'}</span>
            </button>

            {/* Column Selector */}
            <button
              onClick={() => setIsColumnModalOpen(true)}
              className="px-3 py-1.5 text-xs font-black rounded-xl bg-slate-800 hover:bg-slate-700 text-indigo-300 border border-indigo-500/30 transition-all flex items-center gap-1.5 cursor-pointer"
              title="ကော်လံများ စိတ်ကြိုက် ဖွင့်/ပိတ်ပါ"
            >
              <Columns size={13} />
              <span>Columns ({visibleColumns.length})</span>
            </button>

            {/* Export Excel Button */}
            <button
              onClick={exportToExcel}
              className="px-3.5 py-2 text-xs font-black rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white shadow-lg transition-all flex items-center gap-1.5 cursor-pointer active:scale-95"
              title="Export to Excel (.xlsx)"
            >
              <FileSpreadsheet size={14} />
              <span>Export Excel</span>
            </button>

            {/* Direct Save PDF Button */}
            <button
              onClick={exportDirectPDF}
              disabled={isExportingPDF}
              className="px-3.5 py-2 text-xs font-black rounded-xl bg-rose-600 hover:bg-rose-500 text-white shadow-lg transition-all flex items-center gap-1.5 cursor-pointer active:scale-95 disabled:opacity-50"
              title="Save Direct A4 Landscape PDF"
            >
              {isExportingPDF ? <Loader2 size={14} className="animate-spin" /> : <FileDown size={14} />}
              <span>Save as PDF</span>
            </button>

            {/* Print Button */}
            <button
              onClick={triggerPrint}
              className="px-3.5 py-2 text-xs font-black rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white shadow-lg transition-all flex items-center gap-1.5 cursor-pointer active:scale-95"
              title="Print Table (A4 Landscape)"
            >
              <Printer size={14} />
              <span>Print Table</span>
            </button>
          </div>
        </div>

        {/* Minimal Signer Control Strip */}
        <div className="pt-4 border-t border-slate-800 flex flex-col md:flex-row gap-3 items-start md:items-center justify-between">
          <button
            type="button"
            onClick={() => handleToggleShowSigner(!showSigner)}
            className={`px-3 py-1.5 rounded-xl font-bold text-xs transition-all flex items-center gap-1.5 cursor-pointer border ${
              showSigner
                ? 'bg-emerald-500/20 text-emerald-300 border-emerald-500/40'
                : 'bg-slate-800 text-slate-400 border-slate-700'
            }`}
          >
            <Shield size={13} className={showSigner ? "text-emerald-400" : "text-slate-500"} />
            <span>{showSigner ? '✓ Signer ပါဝင်မည်' : '✕ Signer ဖြုတ်ထားသည်'}</span>
          </button>

          {showSigner && (
            <div className="flex flex-wrap items-center gap-3 w-full md:w-auto">
              <div className="flex items-center gap-2 bg-slate-800/90 border border-slate-700 px-3 py-1.5 rounded-xl flex-1 md:flex-initial">
                <span className="text-[10px] text-slate-400 font-bold uppercase shrink-0">ရာထူး:</span>
                <input
                  type="text"
                  value={officerTitle}
                  onChange={(e) => handleSaveOfficerTitle(e.target.value)}
                  list="crtOfficialTitlesList"
                  placeholder="ရာထူး..."
                  className="bg-transparent text-white font-bold text-xs outline-none font-pyidaungsu w-32"
                />
              </div>

              <div className="flex items-center gap-2 bg-slate-800/90 border border-slate-700 px-3 py-1.5 rounded-xl flex-1 md:flex-initial">
                <span className="text-[10px] text-slate-400 font-bold uppercase shrink-0">အမည်:</span>
                <input
                  type="text"
                  value={officerName}
                  onChange={(e) => handleSaveOfficerName(e.target.value)}
                  list="crtOfficialNamesList"
                  placeholder="အရာရှိ အမည်..."
                  className="bg-transparent text-white font-bold text-xs outline-none font-pyidaungsu w-36"
                />
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Datalists for Custom Signer */}
      <datalist id="crtOfficialTitlesList">
        <option value="ဒုတိယလဝကမှူး" />
        <option value="လဝကမှူး" />
        <option value="ဦးစီးအရာရှိ" />
        <option value="လက်ထောက်ညွှန်ကြားရေးမှူး" />
        <option value="ဒုတိယညွှန်ကြားရေးမှူး" />
        <option value="ညွှန်ကြားရေးမှူး" />
        <option value="တာဝန်ခံအရာရှိ" />
        {masterData.filter(m => m.type === 'Title' && m.name).map(m => (
          <option key={`title-${m.id}`} value={m.name} />
        ))}
      </datalist>
      <datalist id="crtOfficialNamesList">
        {currentUser?.name && <option value={currentUser.name} />}
        {masterData.filter(m => m.type === 'Official' && m.name).map(m => (
          <option key={`off-${m.id}`} value={m.name} />
        ))}
      </datalist>

      {/* Direct On-Screen Filter Toolbar (No Sub-Tabs / Hidden Panels) */}
      <div className="bg-white rounded-3xl p-5 shadow-md border border-slate-200/80 space-y-4 no-print">
        {/* Top Row: Data Pool Selector & Counts */}
        <div className="flex flex-wrap items-center justify-between gap-3 pb-3 border-b border-slate-100">
          {/* Data Pool Selector */}
          <div className="flex items-center gap-1 bg-slate-100 p-1 rounded-2xl">
            <button
              onClick={() => setDataPool('STILL_IN')}
              className={`px-3 py-1.5 rounded-xl font-bold text-xs transition-all cursor-pointer ${
                dataPool === 'STILL_IN'
                  ? 'bg-indigo-600 text-white shadow-sm'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              လက်ရှိနေထိုင်ဆဲ (Still In)
            </button>
            <button
              onClick={() => setDataPool('ALL_MOVEMENTS')}
              className={`px-3 py-1.5 rounded-xl font-bold text-xs transition-all cursor-pointer ${
                dataPool === 'ALL_MOVEMENTS'
                  ? 'bg-slate-800 text-white shadow-sm'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              အားလုံး (All Records)
            </button>
            <button
              onClick={() => setDataPool('INBOUND')}
              className={`px-3 py-1.5 rounded-xl font-bold text-xs transition-all cursor-pointer ${
                dataPool === 'INBOUND'
                  ? 'bg-emerald-700 text-white shadow-sm'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              ဝင်ရောက်သူများ (Inbound)
            </button>
            <button
              onClick={() => setDataPool('OUTBOUND')}
              className={`px-3 py-1.5 rounded-xl font-bold text-xs transition-all cursor-pointer ${
                dataPool === 'OUTBOUND'
                  ? 'bg-orange-600 text-white shadow-sm'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              ထွက်ခွါသူများ (Outbound)
            </button>
          </div>

          {/* Quick Selection Actions & Count Indicator */}
          <div className="flex items-center gap-2">
            <span className="text-xs font-bold text-slate-500">
              ပါဝင်သူ: <strong className="text-indigo-900 font-black">{displayRows.length}</strong> / {sortedRows.length} ဦး
            </span>
            <button
              onClick={() => toggleSelectAll(true)}
              className="px-2.5 py-1 text-xs font-bold rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 flex items-center gap-1 cursor-pointer transition-colors"
            >
              <CheckSquare size={13} />
              <span>အားလုံးရွေး</span>
            </button>
            <button
              onClick={() => toggleSelectAll(false)}
              className="px-2.5 py-1 text-xs font-bold rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 flex items-center gap-1 cursor-pointer transition-colors"
            >
              <Square size={13} />
              <span>ပယ်ဖျက်</span>
            </button>
            <button
              onClick={() => {
                setSearchQuery('');
                setFilterNationality('ALL');
                setFilterVisaType('ALL');
                setFilterAddress('ALL');
                setFilterGender('ALL');
                setFilterAddressType('ALL');
                setExcludeVisaS(false);
                setMinElapsedDays('');
                setStartDate('');
                setEndDate('');
                setSortBy('elapsed');
                setSortDirection('desc');
                showToast("Filters Reset ပြုလုပ်ပြီးပါပြီ");
              }}
              className="p-1.5 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-lg cursor-pointer transition-colors"
              title="Reset All Filters"
            >
              <RotateCcw size={13} />
            </button>
          </div>
        </div>

        {/* Direct Filter Controls Grid */}
        <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3">
          {/* Live Search */}
          <div className="lg:col-span-2 relative">
            <Search size={14} className="absolute left-3 top-2.5 text-slate-400" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="ရှာဖွေရန် (ပတ်စပို့၊ အမည်၊ လိပ်စာ)..."
              className="w-full pl-8 pr-3 py-1.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:bg-white focus:border-indigo-500 outline-none"
            />
          </div>

          {/* Nationality Filter */}
          <div>
            <select
              value={filterNationality}
              onChange={(e) => setFilterNationality(e.target.value)}
              className="w-full py-1.5 px-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 outline-none cursor-pointer"
            >
              <option value="ALL">နိုင်ငံသား အားလုံး</option>
              {uniqueNationalities.map(n => (
                <option key={n} value={n}>{n}</option>
              ))}
            </select>
          </div>

          {/* Visa Type Filter */}
          <div>
            <select
              value={filterVisaType}
              onChange={(e) => setFilterVisaType(e.target.value)}
              className="w-full py-1.5 px-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 outline-none cursor-pointer"
            >
              <option value="ALL">ဗီဇာ အားလုံး</option>
              {uniqueVisaTypes.map(v => (
                <option key={v} value={v}>{v}</option>
              ))}
            </select>
          </div>

          {/* Address Category */}
          <div>
            <select
              value={filterAddressType}
              onChange={(e) => setFilterAddressType(e.target.value as any)}
              className="w-full py-1.5 px-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 outline-none cursor-pointer"
            >
              <option value="ALL">လိပ်စာ အားလုံး</option>
              <option value="OTHER">ဟိုတယ် / အခြားလိပ်စာ</option>
              <option value="COMPANY">ကုမ္ပဏီ (Company)</option>
            </select>
          </div>

          {/* Sorting Field */}
          <div className="flex items-center gap-1.5">
            <select
              value={sortBy}
              onChange={(e) => setSortBy(e.target.value as any)}
              className="w-full py-1.5 px-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-indigo-900 outline-none cursor-pointer"
            >
              <option value="elapsed">ရက်ပေါင်း (Elapsed)</option>
              <option value="arrival">ရောက်ရှိရက် (Arrival)</option>
              <option value="passport">ပတ်စပို့ (Passport)</option>
              <option value="name">အမည် (Name)</option>
              <option value="nationality">နိုင်ငံ (Nationality)</option>
              <option value="address">လိပ်စာ (Address)</option>
              <option value="visa">ဗီဇာ (Visa)</option>
              <option value="stayTo">သက်တမ်းကုန် (Stay Expiry)</option>
              <option value="custom">စိတ်ကြိုက် (Custom)</option>
            </select>

            <button
              onClick={() => setSortDirection(prev => prev === 'asc' ? 'desc' : 'asc')}
              className="p-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl cursor-pointer border border-slate-200 shrink-0"
              title={sortDirection === 'asc' ? 'ငယ်စဉ်ကြီးလိုက် (Asc)' : 'ကြီးစဉ်ငယ်လိုက် (Desc)'}
            >
              <ArrowUpDown size={14} />
            </button>
          </div>
        </div>

        {/* Secondary Filters Strip (Visa S Exclude Toggle, Gender, Min Days & Date Range) */}
        <div className="flex flex-wrap items-center gap-3 pt-2 border-t border-slate-100 text-xs">
          {/* Exclude Visa S Toggle */}
          <label className="flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-amber-50 border border-amber-200 text-amber-900 font-bold cursor-pointer select-none">
            <input
              type="checkbox"
              checked={excludeVisaS}
              onChange={(e) => setExcludeVisaS(e.target.checked)}
              className="w-3.5 h-3.5 rounded text-amber-600 cursor-pointer"
            />
            <span>ဗီဇာ အမျိုးအစား S မပါ (Exclude Special/Stay)</span>
          </label>

          {/* Gender */}
          <div className="flex items-center gap-1.5">
            <span className="text-slate-400 font-bold">ကျား/မ:</span>
            <select
              value={filterGender}
              onChange={(e) => setFilterGender(e.target.value as any)}
              className="py-1 px-2 bg-slate-50 border border-slate-200 rounded-lg font-bold text-slate-700 outline-none cursor-pointer text-xs"
            >
              <option value="ALL">အားလုံး</option>
              <option value="M">ကျား (M)</option>
              <option value="F">မ (F)</option>
            </select>
          </div>

          {/* Min Days */}
          <div className="flex items-center gap-1.5">
            <span className="text-slate-400 font-bold">အနည်းဆုံး ရက်:</span>
            <input
              type="number"
              value={minElapsedDays}
              onChange={(e) => setMinElapsedDays(e.target.value)}
              placeholder="ရက်ပေါင်း..."
              className="w-20 py-1 px-2 bg-slate-50 border border-slate-200 rounded-lg font-bold text-slate-700 outline-none text-xs"
            />
          </div>

          {/* Date Range */}
          <div className="flex items-center gap-1.5">
            <span className="text-slate-400 font-bold">ရောက်ရှိရက်စွဲ:</span>
            <input
              type="date"
              value={startDate}
              onChange={(e) => setStartDate(e.target.value)}
              className="py-1 px-2 bg-slate-50 border border-slate-200 rounded-lg font-bold text-slate-700 outline-none text-xs"
            />
            <span className="text-slate-400 font-bold">မှ</span>
            <input
              type="date"
              value={endDate}
              onChange={(e) => setEndDate(e.target.value)}
              className="py-1 px-2 bg-slate-50 border border-slate-200 rounded-lg font-bold text-slate-700 outline-none text-xs"
            />
          </div>
        </div>
      </div>

      {/* Row Selection & Custom Sorting Drawer */}
      <div className="bg-slate-50 border border-slate-200 rounded-3xl p-4 no-print">
        <div className="flex items-center justify-between mb-2">
          <span className="text-xs font-black uppercase text-slate-700 tracking-wider">
            Row Selection & Custom Sequence ({sortedRows.length} items)
          </span>
          <span className="text-[10px] text-slate-500 font-semibold hidden sm:inline">
            ဇယားတွင် ပါဝင်လိုသူများကို Checkbox ဖြင့် ရွေးချယ်နိုင်ပြီး ▲ / ▼ ဖြင့် အစီအစဉ် ရွှေ့နိုင်ပါသည်
          </span>
        </div>

        <div className="max-h-48 overflow-y-auto bg-white rounded-2xl border border-slate-200 divide-y divide-slate-100">
          {sortedRows.length === 0 ? (
            <div className="p-4 text-center text-xs text-slate-400 italic">
              စစ်ထုတ်ထားသော အချက်အလက် မရှိပါ။
            </div>
          ) : (
            sortedRows.map((r, idx) => (
              <div 
                key={r.id} 
                className={`p-2 flex items-center justify-between text-xs transition-colors ${
                  r.selected ? 'bg-indigo-50/30' : 'bg-slate-50/50 opacity-60'
                }`}
              >
                <div className="flex items-center gap-2.5 flex-1 min-w-0">
                  <input
                    type="checkbox"
                    checked={r.selected}
                    onChange={() => toggleRowSelect(r.id)}
                    className="w-3.5 h-3.5 rounded text-indigo-600 cursor-pointer"
                  />
                  <span className="font-bold text-slate-400 w-5 text-center text-[10px]">{idx + 1}</span>
                  <span className="font-mono font-black text-indigo-950 text-[11px]">{r.passport}</span>
                  <span className="font-bold text-slate-800 truncate max-w-[160px]">{r.fullname}</span>
                  <span className="text-[10px] bg-slate-100 text-slate-600 px-1.5 py-0.2 rounded font-bold">{r.nationality}</span>
                  <span className="text-[10px] text-slate-500 truncate max-w-[180px]">{r.address}</span>
                  <span className="text-[10px] font-bold text-indigo-700">{r.elapsedDays} ရက်</span>
                </div>

                <div className="flex items-center gap-1">
                  <button
                    onClick={() => handleMoveRow(r.id, 'up')}
                    disabled={idx === 0}
                    className="p-1 rounded bg-slate-100 hover:bg-slate-200 text-slate-600 disabled:opacity-30 cursor-pointer"
                    title="အပေါ်သို့ ရွှေ့မည်"
                  >
                    <ArrowUp size={12} />
                  </button>
                  <button
                    onClick={() => handleMoveRow(r.id, 'down')}
                    disabled={idx === sortedRows.length - 1}
                    className="p-1 rounded bg-slate-100 hover:bg-slate-200 text-slate-600 disabled:opacity-30 cursor-pointer"
                    title="အောက်သို့ ရွှေ့မည်"
                  >
                    <ArrowDown size={12} />
                  </button>
                </div>
              </div>
            ))
          )}
        </div>
      </div>

      {/* The Printable A4 Landscape Report Sheet View */}
      <div 
        ref={previewContainerRef}
        className="bg-slate-100/90 rounded-3xl p-3 sm:p-6 border border-slate-200/80 w-full max-w-full overflow-hidden flex flex-col items-center shadow-inner"
      >
        <div className="w-full max-w-[297mm] flex justify-between items-center gap-3 mb-3 no-print">
          <span className="text-xs font-black uppercase text-slate-700 tracking-wider">
            Official Report Sheet (A4 Landscape Preview)
          </span>

          {/* View Zoom Mode Switcher */}
          <div className="flex items-center bg-white border border-slate-200 rounded-xl p-0.5 shadow-xs">
            <button
              type="button"
              onClick={() => setViewZoomMode('fit')}
              className={`px-2.5 py-1 rounded-lg text-[11px] font-black transition-all flex items-center gap-1 cursor-pointer ${
                viewZoomMode === 'fit'
                  ? 'bg-indigo-600 text-white shadow-xs'
                  : 'text-slate-600 hover:text-indigo-600 hover:bg-slate-50'
              }`}
              title="Fit Screen"
            >
              <Minimize2 size={12} />
              <span>Fit Screen {fitScale < 1 && `(${Math.round(fitScale * 100)}%)`}</span>
            </button>
            <button
              type="button"
              onClick={() => setViewZoomMode('full')}
              className={`px-2.5 py-1 rounded-lg text-[11px] font-black transition-all flex items-center gap-1 cursor-pointer ${
                viewZoomMode === 'full'
                  ? 'bg-indigo-600 text-white shadow-xs'
                  : 'text-slate-600 hover:text-indigo-600 hover:bg-slate-50'
              }`}
              title="100% Full Size"
            >
              <Maximize2 size={12} />
              <span>100% Full Size</span>
            </button>
          </div>
        </div>

        {/* Paper Container Wrapper */}
        <div
          className="w-full flex flex-col items-center gap-6"
          style={
            viewZoomMode === 'fit' && currentScale < 1
              ? {
                  width: '100%',
                  overflow: 'hidden',
                }
              : {
                  width: '100%',
                  overflowX: 'auto',
                  WebkitOverflowScrolling: 'touch',
                }
          }
        >
          <div
            ref={printRef}
            style={
              viewZoomMode === 'fit' && currentScale < 1
                ? {
                    width: '297mm',
                    minWidth: '297mm',
                    transform: `scale(${currentScale})`,
                    transformOrigin: 'top center',
                  }
                : {
                    width: '297mm',
                    minWidth: '297mm',
                  }
            }
            className="transition-transform duration-150 origin-top flex flex-col items-center gap-8 print:gap-0"
          >
            {paginatedPages.map((pageRows, pageIdx) => {
              const isLastPage = pageIdx === paginatedPages.length - 1;
              const totalPages = paginatedPages.length;

              return (
                <div
                  key={`crt-page-${pageIdx}`}
                  data-page-index={pageIdx}
                  className="crt-pdf-page bg-white text-black p-[12mm] shadow-2xl rounded-sm w-[297mm] min-w-[297mm] max-w-[297mm] min-h-[210mm] max-h-[210mm] h-[210mm] border border-slate-300 relative font-pyidaungsu text-[12px] select-text box-border flex flex-col justify-between print:border-none print:shadow-none print:break-after-page print:mb-0"
                >
                  <div>
                    {/* Customizable Title Header (Non-bold, clean Pyidaungsu) */}
                    <div className="text-center font-normal text-[15px] text-black font-pyidaungsu mb-1 tracking-normal">
                      {tableTitle}
                    </div>
                    <div className="text-right font-normal text-[12px] text-black font-pyidaungsu mb-2.5 pr-1">
                      ရက်စွဲ၊ {useBurmeseDigits ? toBurmeseSlashDate(selectedDate) : formatToDDMMYYYY(selectedDate)}
                    </div>

                    {/* Data Table with Repeating Header on Every Page */}
                    <table className="w-full text-center text-[11px] border-collapse border border-black font-pyidaungsu table-fixed">
                      <colgroup>
                        {CRT_COLUMNS_DEFINITION.filter(col => visibleColumns.includes(col.id)).map(col => (
                          <col key={col.id} className={col.defaultColWidth} />
                        ))}
                      </colgroup>
                      <thead>
                        <tr className="border-b border-black align-middle text-center bg-slate-50 font-normal">
                          {visibleColumns.includes('sr') && (
                            <th 
                              onClick={() => handleHeaderSort('custom')} 
                              className="border border-black p-1.5 font-normal text-[11px] text-black cursor-pointer hover:bg-slate-200 transition-colors"
                            >
                              စဉ်
                            </th>
                          )}
                          {visibleColumns.includes('passport') && (
                            <th 
                              onClick={() => handleHeaderSort('passport')} 
                              className="border border-black p-1.5 font-normal text-[11px] text-black cursor-pointer hover:bg-slate-200 transition-colors"
                            >
                              နိုင်ငံကူးလက်မှတ်အမှတ်
                            </th>
                          )}
                          {visibleColumns.includes('nationality') && (
                            <th 
                              onClick={() => handleHeaderSort('nationality')} 
                              className="border border-black p-1.5 font-normal text-[11px] text-black cursor-pointer hover:bg-slate-200 transition-colors"
                            >
                              နိုင်ငံအမည်
                            </th>
                          )}
                          {visibleColumns.includes('fullname') && (
                            <th 
                              onClick={() => handleHeaderSort('name')} 
                              className="border border-black p-1.5 font-normal text-[11px] text-black cursor-pointer hover:bg-slate-200 transition-colors"
                            >
                              အမည်
                            </th>
                          )}
                          {visibleColumns.includes('gender') && (
                            <th className="border border-black p-1.5 font-normal text-[11px] text-black">
                              ကျား/မ
                            </th>
                          )}
                          {visibleColumns.includes('visaType') && (
                            <th 
                              onClick={() => handleHeaderSort('visa')} 
                              className="border border-black p-1.5 font-normal text-[11px] text-black cursor-pointer hover:bg-slate-200 transition-colors"
                            >
                              ဗီဇာအမျိုးအစား
                            </th>
                          )}
                          {visibleColumns.includes('stayPeriod') && (
                            <th 
                              onClick={() => handleHeaderSort('stayTo')} 
                              className="border border-black p-1.5 font-normal text-[11px] text-black cursor-pointer hover:bg-slate-200 transition-colors"
                            >
                              ဗီဇာသက်တမ်း (မှ/ထိ)
                            </th>
                          )}
                          {visibleColumns.includes('arrivalDate') && (
                            <th 
                              onClick={() => handleHeaderSort('arrival')} 
                              className="border border-black p-1.5 font-normal text-[11px] text-black cursor-pointer hover:bg-slate-200 transition-colors"
                            >
                              နောက်ဆုံးရောက်ရှိရက်စွဲ
                            </th>
                          )}
                          {visibleColumns.includes('elapsedDays') && (
                            <th 
                              onClick={() => handleHeaderSort('elapsed')} 
                              className="border border-black p-1.5 font-normal text-[11px] text-black cursor-pointer hover:bg-slate-200 transition-colors"
                            >
                              နောက်ဆုံးရောက်ရှိသည့်နေ့မှ ရက်ပေါင်း
                            </th>
                          )}
                          {visibleColumns.includes('address') && (
                            <th 
                              onClick={() => handleHeaderSort('address')} 
                              className="border border-black p-1.5 font-normal text-[11px] text-black cursor-pointer hover:bg-slate-200 transition-colors"
                            >
                              တည်းခိုလိပ်စာ
                            </th>
                          )}
                          {visibleColumns.includes('remarks') && (
                            <th className="border border-black p-1.5 font-normal text-[11px] text-black">
                              မှတ်ချက်
                            </th>
                          )}
                        </tr>
                      </thead>
                      <tbody>
                        {pageRows.length === 0 ? (
                          <tr>
                            <td colSpan={visibleColumns.length || 1} className="p-6 text-center text-slate-400 font-normal italic border border-black">
                              အချက်အလက် စာရင်းမရှိပါ။
                            </td>
                          </tr>
                        ) : (
                          pageRows.map((m, rowInPageIdx) => {
                            const overallIdx = pageIdx * ROWS_PER_PAGE + rowInPageIdx;
                            return (
                              <tr key={m.id} className="border-b border-black font-normal">
                                {visibleColumns.includes('sr') && (
                                  <td className="p-1.5 border border-black text-center font-normal">
                                    {useBurmeseDigits ? toBurmeseDigits(overallIdx + 1) : overallIdx + 1}
                                  </td>
                                )}
                                {visibleColumns.includes('passport') && (
                                  <td className="p-1.5 border border-black text-center font-normal font-mono break-words">
                                    {m.passport}
                                  </td>
                                )}
                                {visibleColumns.includes('nationality') && (
                                  <td className="p-1.5 border border-black text-center font-normal break-words">
                                    {m.nationality}
                                  </td>
                                )}
                                {visibleColumns.includes('fullname') && (
                                  <td className="p-1.5 border border-black text-left font-normal break-words">
                                    {m.fullname}
                                  </td>
                                )}
                                {visibleColumns.includes('gender') && (
                                  <td className="p-1.5 border border-black text-center font-normal">
                                    {m.gender === 'M' ? 'ကျား' : m.gender === 'F' ? 'မ' : '-'}
                                  </td>
                                )}
                                {visibleColumns.includes('visaType') && (
                                  <td className="p-1.5 border border-black text-center font-normal break-words">
                                    {m.visaType}
                                  </td>
                                )}
                                {visibleColumns.includes('stayPeriod') && (
                                  <td className="p-1.5 border border-black text-center font-normal break-words">
                                    {(m.stayFrom || m.stayTo) ? `${m.stayFrom || '-'} မှ ${m.stayTo || '-'}` : '-'}
                                  </td>
                                )}
                                {visibleColumns.includes('arrivalDate') && (
                                  <td className="p-1.5 border border-black text-center font-normal">
                                    {m.lastArrivalDate}
                                  </td>
                                )}
                                {visibleColumns.includes('elapsedDays') && (
                                  <td className="p-1.5 border border-black text-center font-normal">
                                    {useBurmeseDigits ? toBurmeseDigits(m.elapsedDays) : m.elapsedDays}
                                  </td>
                                )}
                                {visibleColumns.includes('address') && (
                                  <td className="p-1.5 border border-black text-left font-normal break-words">
                                    {m.address}
                                  </td>
                                )}
                                {visibleColumns.includes('remarks') && (
                                  <td className="p-1.5 border border-black text-left font-normal break-words text-[10px]">
                                    {m.remarks || ''}
                                  </td>
                                )}
                              </tr>
                            );
                          })
                        )}
                      </tbody>
                    </table>
                  </div>

                  {/* Footer: Signer on last page, and page count */}
                  <div className="mt-auto pt-2">
                    {isLastPage && showSigner && (
                      <div className="flex justify-end pr-6 mb-2">
                        <div className="text-center font-pyidaungsu flex flex-col items-center min-w-[260px] group">
                          <div className="h-8" />
                          <div className="w-full">
                            <input
                              type="text"
                              value={officerTitle}
                              onChange={(e) => handleSaveOfficerTitle(e.target.value)}
                              placeholder="ရာထူး..."
                              list="crtOfficialTitlesList"
                              className="w-full text-center text-[12px] font-normal text-black bg-transparent border-b border-transparent group-hover:border-slate-300 focus:border-indigo-500 focus:bg-white outline-none py-0.5 transition-all font-pyidaungsu"
                            />
                          </div>
                          <div className="w-full mt-0.5">
                            <input
                              type="text"
                              value={officerName}
                              onChange={(e) => handleSaveOfficerName(e.target.value)}
                              placeholder="အရာရှိ အမည်..."
                              list="crtOfficialNamesList"
                              className="w-full text-center text-[12px] font-normal text-black bg-transparent border-b border-transparent group-hover:border-slate-300 focus:border-indigo-500 focus:bg-white outline-none py-0.5 transition-all font-pyidaungsu"
                            />
                          </div>
                        </div>
                      </div>
                    )}

                    {totalPages > 1 && (
                      <div className="text-center text-[10px] text-slate-500 font-normal font-pyidaungsu">
                        စာမျက်နှာ {useBurmeseDigits ? toBurmeseDigits(pageIdx + 1) : pageIdx + 1} / {useBurmeseDigits ? toBurmeseDigits(totalPages) : totalPages}
                      </div>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* Column Visibility Modal */}
      <CustomReportColumnModal
        isOpen={isColumnModalOpen}
        onClose={() => setIsColumnModalOpen(false)}
        visibleColumns={visibleColumns}
        onToggleColumn={handleToggleColumn}
        onSelectAll={handleSelectAllColumns}
        onResetDefault={handleResetDefaultColumns}
      />

      <style>{`
        @media print {
          @page {
            size: A4 landscape;
            margin: 0;
          }
          body {
            background-color: #ffffff !important;
            color: #000000 !important;
            margin: 0 !important;
            padding: 0 !important;
          }
          .no-print {
            display: none !important;
          }
          .crt-pdf-page {
            page-break-after: always !important;
            break-after: page !important;
            margin: 0 !important;
            border: none !important;
            box-shadow: none !important;
            width: 297mm !important;
            height: 210mm !important;
          }
        }
      `}</style>
    </div>
  );
};
