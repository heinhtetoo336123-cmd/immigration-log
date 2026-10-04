import React, { useState, useMemo, useRef, useEffect } from 'react';
import { 
  FileSpreadsheet, Printer, RotateCcw, 
  Search, CheckSquare, Square, ArrowUpDown, ArrowUp, ArrowDown,
  Edit3, Check, Shield, Columns, Minimize2, Maximize2,
  Building2, FileDown, Loader2, Share2, ChevronDown, ChevronUp,
  Filter, X, SlidersHorizontal
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import html2canvas from 'html2canvas';
import { ImmRecord, MovementData, MasterItem, CRTColumnId } from '../types';
import { CRT_COLUMNS_DEFINITION, ALL_COLUMN_IDS } from '../utils/crtPresets';
import { CustomReportColumnModal } from './CustomReportColumnModal';
import { logActivity } from '../utils/activityLogger';
import { formatToDDMMYYYY, normalizeStandardDate } from '../App';

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
  const [isGeneratingPDF, setIsGeneratingPDF] = useState<boolean>(false);
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

  // Collapsible Sections States (Default Collapsed for Executive Dashboard)
  const [isFilterExpanded, setIsFilterExpanded] = useState<boolean>(false);
  const [isChecklistExpanded, setIsChecklistExpanded] = useState<boolean>(false);

  // Active Advanced Filters Count Badge
  const activeAdvancedFilterCount = useMemo(() => {
    let count = 0;
    if (filterNationality !== 'ALL') count++;
    if (filterVisaType !== 'ALL') count++;
    if (filterAddress !== 'ALL') count++;
    if (filterGender !== 'ALL') count++;
    if (filterAddressType !== 'ALL') count++;
    if (excludeVisaS) count++;
    if (minElapsedDays) count++;
    if (startDate || endDate) count++;
    return count;
  }, [filterNationality, filterVisaType, filterAddress, filterGender, filterAddressType, excludeVisaS, minElapsedDays, startDate, endDate]);

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

  // Fast & Robust PDF Blob Generator (100% Isolated inside a Detached Hidden Iframe)
  const generatePDFBlob = async (): Promise<{ blob: Blob; fileName: string } | null> => {
    const selectedReportDate = normalizeStandardDate(selectedDate);
    const fileName = `လက်ရှိ_ဟိုတယ်နေနိုင်ငံခြားသားစာရင်း_${selectedReportDate}.pdf`;

    const formattedDate = useBurmeseDigits ? toBurmeseDigits(normalizeStandardDate(selectedDate)) : normalizeStandardDate(selectedDate);
    const columnsHtml = visibleColumns.map(colId => {
      const col = CRT_COLUMNS_DEFINITION.find(c => c.id === colId);
      return `<th style="border: 1px solid #000000; padding: 6px 4px; font-weight: normal; font-size: 11px; background-color: #f8fafc; font-family: 'Pyidaungsu', Myanmar3, sans-serif; text-align: center; color: #000000;">${col?.label || ''}</th>`;
    }).join('');

    const rowsHtml = displayRows.map((r, idx) => {
      const sr = useBurmeseDigits ? toBurmeseDigits(idx + 1) : (idx + 1);
      const elapsed = useBurmeseDigits ? toBurmeseDigits(r.elapsedDays) : r.elapsedDays;
      const gender = r.gender === 'M' ? 'ကျား' : r.gender === 'F' ? 'မ' : '-';
      const visaPeriod = (r.stayFrom || r.stayTo) ? `${normalizeStandardDate(r.stayFrom)} မှ ${normalizeStandardDate(r.stayTo)}` : '-';
      const arrivalDate = r.lastArrivalDate ? (useBurmeseDigits ? toBurmeseDigits(normalizeStandardDate(r.lastArrivalDate)) : normalizeStandardDate(r.lastArrivalDate)) : '-';

      let cells = '';
      if (visibleColumns.includes('sr')) cells += `<td style="border: 1px solid #000000; padding: 5px 3px; text-align: center; font-size: 11px; color: #000000;">${sr}</td>`;
      if (visibleColumns.includes('passport')) cells += `<td style="border: 1px solid #000000; padding: 5px 3px; text-align: center; font-family: monospace; font-size: 11px; color: #000000;">${r.passport}</td>`;
      if (visibleColumns.includes('nationality')) cells += `<td style="border: 1px solid #000000; padding: 5px 3px; text-align: center; font-size: 11px; color: #000000;">${r.nationality}</td>`;
      if (visibleColumns.includes('fullname')) cells += `<td style="border: 1px solid #000000; padding: 5px 4px; text-align: left; font-size: 11px; color: #000000;">${r.fullname}</td>`;
      if (visibleColumns.includes('gender')) cells += `<td style="border: 1px solid #000000; padding: 5px 3px; text-align: center; font-size: 11px; color: #000000;">${gender}</td>`;
      if (visibleColumns.includes('visaType')) cells += `<td style="border: 1px solid #000000; padding: 5px 3px; text-align: center; font-size: 11px; color: #000000;">${r.visaType}</td>`;
      if (visibleColumns.includes('stayPeriod')) cells += `<td style="border: 1px solid #000000; padding: 5px 3px; text-align: center; font-size: 11px; color: #000000;">${visaPeriod}</td>`;
      if (visibleColumns.includes('arrivalDate')) cells += `<td style="border: 1px solid #000000; padding: 5px 3px; text-align: center; font-size: 11px; color: #000000;">${arrivalDate}</td>`;
      if (visibleColumns.includes('elapsedDays')) cells += `<td style="border: 1px solid #000000; padding: 5px 3px; text-align: center; font-size: 11px; color: #000000;">${elapsed}</td>`;
      if (visibleColumns.includes('address')) cells += `<td style="border: 1px solid #000000; padding: 5px 4px; text-align: left; font-size: 11px; color: #000000;">${r.address}</td>`;
      if (visibleColumns.includes('remarks')) cells += `<td style="border: 1px solid #000000; padding: 5px 3px; text-align: left; font-size: 10px; color: #000000;">${r.remarks || ''}</td>`;

      return `<tr>${cells}</tr>`;
    }).join('');

    const signerHtml = showSigner ? `
      <div style="display: flex; justify-content: flex-end; margin-top: 30px; padding-right: 25px;">
        <div style="text-align: center; min-width: 240px; font-family: 'Pyidaungsu', Myanmar3, sans-serif;">
          <div style="height: 30px;"></div>
          <div style="font-size: 12px; font-weight: normal; color: #000000;">(${officerTitle})</div>
          <div style="font-size: 12px; font-weight: normal; color: #000000; margin-top: 3px;">${officerName}</div>
        </div>
      </div>
    ` : '';

    // Create an isolated off-screen hidden iframe - NEVER modify document.body styles or main DOM!
    const iframe = document.createElement('iframe');
    iframe.style.position = 'fixed';
    iframe.style.left = '-9999px';
    iframe.style.top = '0';
    iframe.style.width = '1122px';
    iframe.style.height = '1600px';
    iframe.style.opacity = '0';
    iframe.style.pointerEvents = 'none';
    iframe.style.border = '0';
    iframe.setAttribute('aria-hidden', 'true');
    document.body.appendChild(iframe);

    try {
      const doc = iframe.contentDocument || iframe.contentWindow?.document;
      if (!doc) {
        throw new Error('Unable to access iframe document for PDF generation');
      }

      doc.open();
      doc.write(`
        <!DOCTYPE html>
        <html>
          <head>
            <meta charset="utf-8">
            <style>
              @font-face {
                font-family: 'Pyidaungsu';
                src: local('Pyidaungsu'), local('Myanmar3'), local('Padauk');
              }
              * {
                box-sizing: border-box;
                margin: 0;
                padding: 0;
              }
              body {
                background-color: #ffffff;
                color: #000000;
                font-family: 'Pyidaungsu', Myanmar3, sans-serif;
                font-size: 11px;
                padding: 30px 40px;
                width: 1122px;
              }
              .title-text {
                font-size: 15px;
                font-weight: normal;
                margin-bottom: 3px;
                text-align: center;
                color: #000000;
              }
              .date-text {
                font-size: 11px;
                font-weight: normal;
                margin-bottom: 8px;
                text-align: right;
                color: #000000;
              }
              table {
                width: 100%;
                border-collapse: collapse;
                background-color: #ffffff;
              }
            </style>
          </head>
          <body>
            <div class="title-text">${tableTitle}</div>
            <div class="date-text">ရက်စွဲ၊ ${formattedDate}</div>
            <table>
              <thead>
                <tr>${columnsHtml}</tr>
              </thead>
              <tbody>
                ${rowsHtml}
              </tbody>
            </table>
            ${signerHtml}
          </body>
        </html>
      `);
      doc.close();

      await new Promise(resolve => setTimeout(resolve, 80));

      const canvas = await html2canvas(doc.body, {
        scale: 1.5,
        useCORS: true,
        logging: false,
        backgroundColor: '#ffffff'
      });

      const pdf = new jsPDF({
        orientation: 'landscape',
        unit: 'mm',
        format: 'a4',
        compress: true
      });

      const imgWidth = 297; // A4 landscape width in mm
      const pageHeight = 210; // A4 landscape height in mm
      const imgHeight = (canvas.height * imgWidth) / canvas.width;
      let heightLeft = imgHeight;
      let position = 0;

      const imgData = canvas.toDataURL('image/jpeg', 0.95);
      pdf.addImage(imgData, 'JPEG', 0, position, imgWidth, imgHeight, undefined, 'FAST');
      heightLeft -= pageHeight;

      while (heightLeft > 0) {
        position -= pageHeight;
        pdf.addPage('a4', 'landscape');
        pdf.addImage(imgData, 'JPEG', 0, position, imgWidth, imgHeight, undefined, 'FAST');
        heightLeft -= pageHeight;
      }

      const pdfBlob = pdf.output('blob');
      return { blob: pdfBlob, fileName };
    } finally {
      if (iframe.parentNode) {
        iframe.parentNode.removeChild(iframe);
      }
    }
  };

  // Share Real PDF File directly to Viber / Telegram with 5-Second Hard Timeout Protection
  const handleShareViber = async () => {
    if (displayRows.length === 0) {
      showToast("မျှဝေရန် စာရင်းမရှိပါ");
      return;
    }

    setIsGeneratingPDF(true);
    showToast("Viber သို့ ပေးပို့ရန် PDF ဖိုင် ပြင်ဆင်နေပါသည်...");

    try {
      // 5-second hard timeout protection
      const timeoutPromise = new Promise<null>((_, reject) => {
        setTimeout(() => reject(new Error('PDF_GENERATION_TIMEOUT')), 5000);
      });

      const result = await Promise.race([generatePDFBlob(), timeoutPromise]);

      if (!result) {
        openPrintWindow();
        return;
      }

      const { blob, fileName } = result;
      const pdfFile = new File([blob], fileName, { type: 'application/pdf' });

      if (typeof navigator !== 'undefined' && navigator.canShare && navigator.canShare({ files: [pdfFile] })) {
        await navigator.share({
          title: 'လက်ရှိ ဟိုတယ်နေနိုင်ငံခြားသားစာရင်း',
          files: [pdfFile]
        });
        showToast("Viber / Telegram သို့ PDF ဖိုင် မျှဝေပြီးပါပြီ");
      } else {
        // Desktop / Fallback: automatically download the PDF file
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        showToast("PDF ဖိုင် ဒေါင်းလုဒ်ဆွဲပြီးပါပြီ။ Viber သို့ drag ဆွဲထည့်ပြီး ပေးပို့နိုင်ပါသည်။");
      }
    } catch (err: any) {
      if (err?.name === 'AbortError') {
        // User closed native share sheet - normal action
        return;
      }
      console.warn("Viber PDF Share fallback:", err);
      showToast("PDF တိုက်ရိုက်ထုတ်ယူမှု ဖွင့်လှစ်ပေးပါသည်");
      openPrintWindow();
    } finally {
      // GUARANTEED STATE RESET: Never hang or lock UI
      setIsGeneratingPDF(false);
    }
  };

  // Dedicated Native Print / PDF Popup Engine (100% Pyidaungsu Compatible, Title/Date/Headers Repeated on Every Page)
  const openPrintWindow = () => {
    if (displayRows.length === 0) {
      showToast("ထုတ်ယူရန် စာရင်းမရှိပါ");
      return;
    }

    const printWindow = window.open('', '_blank', 'width=1200,height=800');
    if (!printWindow) {
      window.print();
      return;
    }

    const formattedDate = useBurmeseDigits ? toBurmeseDigits(normalizeStandardDate(selectedDate)) : normalizeStandardDate(selectedDate);
    const numColumns = visibleColumns.length || 1;

    const columnsHtml = visibleColumns.map(colId => {
      const col = CRT_COLUMNS_DEFINITION.find(c => c.id === colId);
      return `<th style="border: 1px solid #000; padding: 6px 4px; font-weight: normal; font-size: 11px; background-color: #f8fafc; font-family: 'Pyidaungsu', Myanmar3, sans-serif; text-align: center;">${col?.label || ''}</th>`;
    }).join('');

    const rowsHtml = displayRows.map((r, idx) => {
      const sr = useBurmeseDigits ? toBurmeseDigits(idx + 1) : (idx + 1);
      const elapsed = useBurmeseDigits ? toBurmeseDigits(r.elapsedDays) : r.elapsedDays;
      const gender = r.gender === 'M' ? 'ကျား' : r.gender === 'F' ? 'မ' : '-';
      const visaPeriod = (r.stayFrom || r.stayTo) ? `${normalizeStandardDate(r.stayFrom)} မှ ${normalizeStandardDate(r.stayTo)}` : '-';
      const arrivalDate = r.lastArrivalDate ? (useBurmeseDigits ? toBurmeseDigits(normalizeStandardDate(r.lastArrivalDate)) : normalizeStandardDate(r.lastArrivalDate)) : '-';

      let cells = '';
      if (visibleColumns.includes('sr')) cells += `<td style="border: 1px solid #000; padding: 5px 3px; text-align: center; font-size: 11px;">${sr}</td>`;
      if (visibleColumns.includes('passport')) cells += `<td style="border: 1px solid #000; padding: 5px 3px; text-align: center; font-family: monospace; font-size: 11px;">${r.passport}</td>`;
      if (visibleColumns.includes('nationality')) cells += `<td style="border: 1px solid #000; padding: 5px 3px; text-align: center; font-size: 11px;">${r.nationality}</td>`;
      if (visibleColumns.includes('fullname')) cells += `<td style="border: 1px solid #000; padding: 5px 4px; text-align: left; font-size: 11px;">${r.fullname}</td>`;
      if (visibleColumns.includes('gender')) cells += `<td style="border: 1px solid #000; padding: 5px 3px; text-align: center; font-size: 11px;">${gender}</td>`;
      if (visibleColumns.includes('visaType')) cells += `<td style="border: 1px solid #000; padding: 5px 3px; text-align: center; font-size: 11px;">${r.visaType}</td>`;
      if (visibleColumns.includes('stayPeriod')) cells += `<td style="border: 1px solid #000; padding: 5px 3px; text-align: center; font-size: 11px;">${visaPeriod}</td>`;
      if (visibleColumns.includes('arrivalDate')) cells += `<td style="border: 1px solid #000; padding: 5px 3px; text-align: center; font-size: 11px;">${arrivalDate}</td>`;
      if (visibleColumns.includes('elapsedDays')) cells += `<td style="border: 1px solid #000; padding: 5px 3px; text-align: center; font-size: 11px;">${elapsed}</td>`;
      if (visibleColumns.includes('address')) cells += `<td style="border: 1px solid #000; padding: 5px 4px; text-align: left; font-size: 11px;">${r.address}</td>`;
      if (visibleColumns.includes('remarks')) cells += `<td style="border: 1px solid #000; padding: 5px 3px; text-align: left; font-size: 10px;">${r.remarks || ''}</td>`;

      return `<tr style="page-break-inside: avoid;">${cells}</tr>`;
    }).join('');

    const signerHtml = showSigner ? `
      <div style="display: flex; justify-content: flex-end; margin-top: 35px; padding-right: 25px; page-break-inside: avoid;">
        <div style="text-align: center; min-width: 240px; font-family: 'Pyidaungsu', Myanmar3, sans-serif;">
          <div style="height: 35px;"></div>
          <div style="font-size: 12px; font-weight: normal; color: #000;">(${officerTitle})</div>
          <div style="font-size: 12px; font-weight: normal; color: #000; margin-top: 3px;">${officerName}</div>
        </div>
      </div>
    ` : '';

    const htmlDoc = `
      <!DOCTYPE html>
      <html>
        <head>
          <meta charset="utf-8">
          <title>${tableTitle} - ${selectedDate}</title>
          <style>
            @import url('https://fonts.googleapis.com/css2?family=Pyidaungsu:wght@400;700&display=swap');
            
            @page {
              size: A4 landscape;
              margin: 10mm;
            }
            * {
              box-sizing: border-box;
              margin: 0;
              padding: 0;
            }
            body {
              font-family: 'Pyidaungsu', Myanmar3, sans-serif;
              font-size: 11px;
              line-height: 1.4;
              color: #000;
              background: #fff;
              padding: 0;
              -webkit-print-color-adjust: exact;
              print-color-adjust: exact;
            }
            table {
              width: 100%;
              border-collapse: collapse;
              table-layout: auto;
            }
            thead {
              display: table-header-group;
            }
            tbody {
              display: table-row-group;
            }
            tr {
              page-break-inside: avoid;
            }
            th, td {
              font-family: 'Pyidaungsu', Myanmar3, sans-serif;
              font-weight: normal;
              color: #000;
            }
            .header-cell {
              border: none !important;
              padding: 0 0 8px 0;
              background: transparent !important;
              text-align: center;
            }
            .title-text {
              font-family: 'Pyidaungsu', Myanmar3, sans-serif;
              font-size: 15px;
              font-weight: normal;
              margin-bottom: 2px;
              color: #000;
              text-align: center;
            }
            .date-text {
              font-family: 'Pyidaungsu', Myanmar3, sans-serif;
              font-size: 11px;
              font-weight: normal;
              color: #000;
              text-align: right;
              padding-right: 4px;
              margin-bottom: 4px;
            }
          </style>
        </head>
        <body>
          <table>
            <thead>
              <tr style="page-break-inside: avoid;">
                <th colspan="${numColumns}" class="header-cell">
                  <div class="title-text">${tableTitle}</div>
                  <div class="date-text">ရက်စွဲ၊ ${formattedDate}</div>
                </th>
              </tr>
              <tr style="page-break-inside: avoid;">${columnsHtml}</tr>
            </thead>
            <tbody>
              ${rowsHtml}
            </tbody>
          </table>
          ${signerHtml}
          <script>
            window.onload = function() {
              setTimeout(function() {
                window.focus();
                window.print();
              }, 300);
            };
          </script>
        </body>
      </html>
    `;

    printWindow.document.open();
    printWindow.document.write(htmlDoc);
    printWindow.document.close();

    logActivity({
      action: 'SETTINGS_CHANGE',
      module: 'CRT',
      targetId: 'Custom_Report_Print',
      details: `Generated Clean Print/PDF (${displayRows.length} rows)`
    });

    showToast("Print / Save as PDF ဖွင့်လှစ်ပြီးပါပြီ");
  };

  return (
    <div className="max-w-7xl mx-auto py-6 px-4 space-y-5">
      {/* 1. Executive Dashboard Header Bar (Row 1) */}
      <div className="bg-slate-900 text-white rounded-3xl p-5 sm:p-6 shadow-xl border border-slate-800 no-print space-y-4">
        <div className="flex flex-col xl:flex-row justify-between items-start xl:items-center gap-4">
          {/* Left: Title, Date, Digits, and Preset */}
          <div className="flex flex-wrap items-center gap-3 flex-1 min-w-0">
            <div className="p-2.5 bg-indigo-600/30 text-indigo-400 rounded-2xl border border-indigo-500/30 shrink-0">
              <FileSpreadsheet size={22} />
            </div>

            {/* Editable Title */}
            <div className="min-w-[200px] max-w-md">
              {isEditingTitle ? (
                <div className="flex items-center gap-2">
                  <input
                    type="text"
                    defaultValue={tableTitle}
                    id="crt_title_input"
                    className="bg-slate-800 text-white border border-indigo-400 px-3 py-1 rounded-xl font-bold text-sm outline-none font-pyidaungsu w-full"
                    autoFocus
                  />
                  <button
                    onClick={() => {
                      const input = document.getElementById('crt_title_input') as HTMLInputElement;
                      if (input) handleSaveTitle(input.value);
                    }}
                    className="p-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl cursor-pointer shrink-0 shadow-md"
                    title="သိမ်းဆည်းမည်"
                  >
                    <Check size={15} />
                  </button>
                </div>
              ) : (
                <div 
                  onClick={() => setIsEditingTitle(true)}
                  className="group cursor-pointer flex items-center gap-2"
                  title="ခေါင်းစဉ်ပြင်ရန် နှိပ်ပါ"
                >
                  <h2 className="text-lg sm:text-xl font-black uppercase tracking-tight font-pyidaungsu text-white group-hover:text-indigo-200 transition-colors truncate">
                    {tableTitle}
                  </h2>
                  <Edit3 size={14} className="text-indigo-400 opacity-60 group-hover:opacity-100 shrink-0" />
                </div>
              )}
            </div>

            {/* Date Input */}
            <div className="flex items-center gap-1.5 bg-slate-800/90 border border-slate-700 px-2.5 py-1.5 rounded-xl">
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
              className={`px-2.5 py-1.5 text-xs font-bold rounded-xl border transition-all cursor-pointer ${
                useBurmeseDigits 
                  ? 'bg-amber-500/20 text-amber-300 border-amber-500/40' 
                  : 'bg-slate-800 text-slate-400 border-slate-700 hover:text-white'
              }`}
              title="နံပါတ် ဂဏန်းပုံစံ ပြောင်းလဲမည်"
            >
              <span>{useBurmeseDigits ? '၁,၂,၃ (မြန်မာ)' : '1,2,3 (Eng)'}</span>
            </button>

            {/* Preset: နေ့စဥ် ဟိုတယ်စာရင်း */}
            <button
              onClick={applyDailyHotelPreset}
              className="px-2.5 py-1.5 text-xs font-bold rounded-xl bg-amber-500/20 hover:bg-amber-500/30 text-amber-300 border border-amber-500/40 transition-all flex items-center gap-1 cursor-pointer"
              title="နေ့စဥ်ပို့ရန် ဟိုတယ်စာရင်း (ဗီဇာ S မပါ) အလိုအလျောက် သတ်မှတ်မည်"
            >
              <Building2 size={13} />
              <span>နေ့စဉ် ဟိုတယ်စာရင်း</span>
            </button>
          </div>

          {/* Right: Clean Uniform Action Buttons Group */}
          <div className="flex flex-wrap items-center gap-2">
            {/* Column Selector */}
            <button
              onClick={() => setIsColumnModalOpen(true)}
              className="h-9 px-3 text-xs font-bold rounded-xl bg-slate-800 hover:bg-slate-700 text-indigo-300 border border-indigo-500/30 transition-all flex items-center gap-1.5 cursor-pointer shadow-sm active:scale-95"
              title="ကော်လံများ စိတ်ကြိုက် ဖွင့်/ပိတ်ပါ"
            >
              <Columns size={14} />
              <span>Columns ({visibleColumns.length})</span>
            </button>

            {/* Export Excel Button */}
            <button
              onClick={exportToExcel}
              className="h-9 px-3.5 text-xs font-bold rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer active:scale-95"
              title="Export to Excel (.xlsx)"
            >
              <FileSpreadsheet size={14} />
              <span>Export Excel</span>
            </button>

            {/* Mobile Share / Viber Button */}
            <button
              onClick={handleShareViber}
              disabled={isGeneratingPDF}
              className="h-9 px-3.5 text-xs font-bold rounded-xl bg-purple-600 hover:bg-purple-500 disabled:opacity-50 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer active:scale-95"
              title="Viber / Telegram သို့ PDF ဖိုင် တိုက်ရိုက် မျှဝေမည်"
            >
              {isGeneratingPDF ? (
                <>
                  <Loader2 size={14} className="animate-spin" />
                  <span>PDF ပြင်ဆင်နေသည်...</span>
                </>
              ) : (
                <>
                  <Share2 size={14} />
                  <span>Share (Viber)</span>
                </>
              )}
            </button>

            {/* Save as PDF / Print Button */}
            <button
              onClick={openPrintWindow}
              className="h-9 px-3.5 text-xs font-bold rounded-xl bg-rose-600 hover:bg-rose-500 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer active:scale-95"
              title="Save as PDF / Print (A4 Landscape Clean Popup)"
            >
              <Printer size={14} />
              <span>Save as PDF / Print</span>
            </button>
          </div>
        </div>

        {/* Minimal Signer Control Strip */}
        <div className="pt-3 border-t border-slate-800/80 flex flex-col md:flex-row gap-2.5 items-start md:items-center justify-between">
          <button
            type="button"
            onClick={() => handleToggleShowSigner(!showSigner)}
            className={`px-2.5 py-1 rounded-xl font-bold text-xs transition-all flex items-center gap-1.5 cursor-pointer border ${
              showSigner
                ? 'bg-emerald-500/20 text-emerald-300 border-emerald-500/40'
                : 'bg-slate-800 text-slate-400 border-slate-700'
            }`}
          >
            <Shield size={12} className={showSigner ? "text-emerald-400" : "text-slate-500"} />
            <span>{showSigner ? '✓ Signer ပါဝင်မည်' : '✕ Signer ဖြုတ်ထားသည်'}</span>
          </button>

          {showSigner && (
            <div className="flex flex-wrap items-center gap-2 w-full md:w-auto">
              <div className="flex items-center gap-1.5 bg-slate-800/90 border border-slate-700 px-2.5 py-1 rounded-xl flex-1 md:flex-initial">
                <span className="text-[10px] text-slate-400 font-bold uppercase shrink-0">ရာထူး:</span>
                <input
                  type="text"
                  value={officerTitle}
                  onChange={(e) => handleSaveOfficerTitle(e.target.value)}
                  list="crtOfficialTitlesList"
                  placeholder="ရာထူး..."
                  className="bg-transparent text-white font-bold text-xs outline-none font-pyidaungsu w-28"
                />
              </div>

              <div className="flex items-center gap-1.5 bg-slate-800/90 border border-slate-700 px-2.5 py-1 rounded-xl flex-1 md:flex-initial">
                <span className="text-[10px] text-slate-400 font-bold uppercase shrink-0">အမည်:</span>
                <input
                  type="text"
                  value={officerName}
                  onChange={(e) => handleSaveOfficerName(e.target.value)}
                  list="crtOfficialNamesList"
                  placeholder="အရာရှိ အမည်..."
                  className="bg-transparent text-white font-bold text-xs outline-none font-pyidaungsu w-32"
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

      {/* 2. Executive Filter Bar (Row 2) */}
      <div className="bg-white rounded-3xl p-4 shadow-sm border border-slate-200/90 space-y-3 no-print">
        {/* Main Line: Segmented Pills, Inline Search, Filters Toggle & Reset */}
        <div className="flex flex-wrap items-center justify-between gap-3">
          {/* Segmented Pill Tabs */}
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

          {/* Inline Search, Advanced Filters Toggle & Reset */}
          <div className="flex items-center gap-2 flex-1 max-w-xl justify-end">
            {/* Inline Search Box with Clear Icon */}
            <div className="relative flex-1 min-w-[180px] max-w-xs">
              <Search size={14} className="absolute left-3 top-2.5 text-slate-400" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="ရှာဖွေရန် (ပတ်စပို့၊ အမည်၊ လိပ်စာ)..."
                className="w-full pl-8 pr-7 py-1.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:bg-white focus:border-indigo-500 outline-none"
              />
              {searchQuery && (
                <button
                  onClick={() => setSearchQuery('')}
                  className="absolute right-2 top-2 text-slate-400 hover:text-slate-600 cursor-pointer"
                  title="Clear Search"
                >
                  <X size={14} />
                </button>
              )}
            </div>

            {/* Expandable Advanced Filters Toggle Button */}
            <button
              onClick={() => setIsFilterExpanded(!isFilterExpanded)}
              className={`px-3 py-1.5 rounded-xl font-bold text-xs transition-all flex items-center gap-1.5 cursor-pointer border ${
                isFilterExpanded || activeAdvancedFilterCount > 0
                  ? 'bg-indigo-50 text-indigo-700 border-indigo-200 shadow-xs'
                  : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
              }`}
              title="အဆင့်မြင့် စစ်ထုတ်မှုများ ဖွင့်/ပိတ်ပါ"
            >
              <SlidersHorizontal size={13} />
              <span>အဆင့်မြင့် စစ်ထုတ်မှုများ {activeAdvancedFilterCount > 0 && `(${activeAdvancedFilterCount})`}</span>
              {isFilterExpanded ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
            </button>

            {/* Reset Filters Button */}
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
              className="p-2 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-xl cursor-pointer transition-colors shrink-0"
              title="Reset All Filters"
            >
              <RotateCcw size={14} />
            </button>
          </div>
        </div>

        {/* Collapsible Advanced Filters Panel (Default Collapsed) */}
        {isFilterExpanded && (
          <div className="pt-3 border-t border-slate-100 space-y-3">
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-2.5">
              {/* Nationality Filter */}
              <div>
                <label className="text-[10px] font-bold text-slate-400 block mb-1">နိုင်ငံသား</label>
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
                <label className="text-[10px] font-bold text-slate-400 block mb-1">ဗီဇာအမျိုးအစား</label>
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
                <label className="text-[10px] font-bold text-slate-400 block mb-1">လိပ်စာအမျိုးအစား</label>
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
              <div>
                <label className="text-[10px] font-bold text-slate-400 block mb-1">အစီအစဉ် စီစဉ်ရန်</label>
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

              {/* Gender */}
              <div>
                <label className="text-[10px] font-bold text-slate-400 block mb-1">ကျား/မ</label>
                <select
                  value={filterGender}
                  onChange={(e) => setFilterGender(e.target.value as any)}
                  className="w-full py-1.5 px-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 outline-none cursor-pointer"
                >
                  <option value="ALL">ကျား/မ အားလုံး</option>
                  <option value="M">ကျား (M)</option>
                  <option value="F">မ (F)</option>
                </select>
              </div>
            </div>

            {/* Extra Row: Exclude Visa S, Min Days, Date Range */}
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
                <span className="text-slate-400 font-bold">ရောက်ရှိရက်:</span>
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
        )}
      </div>

      {/* 3. Row Selection & Custom Sequence Collapsible Bar (Row 3) */}
      <div className="bg-slate-50 border border-slate-200 rounded-2xl p-3 no-print">
        <div className="flex flex-wrap items-center justify-between gap-2">
          {/* Count summary */}
          <div className="flex items-center gap-2">
            <span className="text-xs font-bold text-slate-600">
              ရွေးချယ်ထားသော စာရင်းများ: <strong className="text-indigo-900 font-black">{displayRows.length}</strong> / {sortedRows.length} ဦး
            </span>
          </div>

          {/* Sticky Quick Actions & Collapse Toggle */}
          <div className="flex items-center gap-2">
            <button
              onClick={() => toggleSelectAll(true)}
              className="px-2.5 py-1 text-xs font-bold rounded-lg bg-white border border-slate-200 hover:bg-slate-100 text-slate-700 flex items-center gap-1 cursor-pointer transition-colors shadow-xs"
            >
              <CheckSquare size={13} />
              <span>အားလုံးရွေး</span>
            </button>
            <button
              onClick={() => toggleSelectAll(false)}
              className="px-2.5 py-1 text-xs font-bold rounded-lg bg-white border border-slate-200 hover:bg-slate-100 text-slate-700 flex items-center gap-1 cursor-pointer transition-colors shadow-xs"
            >
              <Square size={13} />
              <span>ပယ်ဖျက်</span>
            </button>
            <button
              onClick={() => setIsChecklistExpanded(!isChecklistExpanded)}
              className="px-3 py-1 text-xs font-bold rounded-lg bg-indigo-50 border border-indigo-200 hover:bg-indigo-100 text-indigo-700 flex items-center gap-1 cursor-pointer transition-colors"
            >
              <span>{isChecklistExpanded ? 'စာရင်း အသေးစိတ် ဝှက်ရန်' : 'စာရင်း အသေးစိတ် ကြည့်ရန် / ပြင်ရန်'}</span>
              {isChecklistExpanded ? <ChevronUp size={13} /> : <ChevronDown size={13} />}
            </button>
          </div>
        </div>

        {/* Expandable Checklist Content (Default Collapsed) */}
        {isChecklistExpanded && (
          <div className="mt-3 pt-3 border-t border-slate-200">
            <div className="max-h-56 overflow-y-auto bg-white rounded-xl border border-slate-200 divide-y divide-slate-100">
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
        )}
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
                                    {(m.stayFrom || m.stayTo) ? (useBurmeseDigits ? toBurmeseDigits(`${normalizeStandardDate(m.stayFrom)} မှ ${normalizeStandardDate(m.stayTo)}`) : `${normalizeStandardDate(m.stayFrom)} မှ ${normalizeStandardDate(m.stayTo)}`) : '-'}
                                  </td>
                                )}
                                {visibleColumns.includes('arrivalDate') && (
                                  <td className="p-1.5 border border-black text-center font-normal">
                                    {m.lastArrivalDate ? (useBurmeseDigits ? toBurmeseDigits(normalizeStandardDate(m.lastArrivalDate)) : normalizeStandardDate(m.lastArrivalDate)) : '-'}
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
