import React, { useState, useMemo, useRef, useEffect } from 'react';
import { 
  FileSpreadsheet, Download, Printer, Filter, ArrowUpDown, 
  Search, CheckSquare, Square, RotateCcw, ArrowUp, ArrowDown, 
  Edit3, Check, Users, MapPin, Globe, Shield, Calendar, ChevronDown, Sparkles,
  Maximize2, Minimize2, Columns, Bookmark
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { ImmRecord, MovementData, MasterItem, Mode, CRTColumnId, CRTPreset } from '../types';
import { CRT_COLUMNS_DEFINITION, ALL_COLUMN_IDS } from '../utils/crtPresets';
import { CustomReportColumnModal } from './CustomReportColumnModal';
import { CustomReportPresetsModal } from './CustomReportPresetsModal';
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
  id: string; // passport or unique key
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
  remarks: string; // User-editable or blank (no auto remark)
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
  const today = new Date().toISOString().split('T')[0];

  // Customizable Table Metadata
  const [tableTitle, setTableTitle] = useState(() => 
    localStorage.getItem('crt_tableTitle') || 'လက်ရှိနေထိုင်သူများမှ အခြားလိပ်စာဖြင့်နေထိုင်သူများ'
  );
  const [isEditingTitle, setIsEditingTitle] = useState(false);
  const [selectedDate, setSelectedDate] = useState<string>(today);
  const [dateRangeMode, setDateRangeMode] = useState<boolean>(false);
  const [startDate, setStartDate] = useState<string>('');
  const [endDate, setEndDate] = useState<string>('');

  // Officer Info
  const [showSigner, setShowSigner] = useState<boolean>(() => {
    const saved = localStorage.getItem('crt_showSigner');
    return saved !== null ? saved === 'true' : true;
  });
  const [officerTitle, setOfficerTitle] = useState(() => 
    localStorage.getItem('crt_officerTitle') || currentUser?.title || 'လဝကမှူး'
  );
  const [officerName, setOfficerName] = useState(() => 
    localStorage.getItem('crt_officerName') || currentUser?.name || 'ဒုတိယလဝကမှူး'
  );

  const handleToggleShowSigner = (val: boolean) => {
    setShowSigner(val);
    try {
      localStorage.setItem('crt_showSigner', String(val));
    } catch {}
    showToast(val ? "Custom Signer လက်မှတ်ပုံစံ ထည့်သွင်းထားပါသည်" : "Custom Signer လက်မှတ်အား ဖြုတ်ထားပါသည်");
  };

  // Digits toggle
  const [useBurmeseDigits, setUseBurmeseDigits] = useState<boolean>(true);

  // Filter States (Multiple Selectable)
  const [dataPool, setDataPool] = useState<'STILL_IN' | 'ALL_MOVEMENTS' | 'INBOUND' | 'OUTBOUND'>('STILL_IN');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [filterNationalities, setFilterNationalities] = useState<string[]>([]);
  const [filterVisaTypes, setFilterVisaTypes] = useState<string[]>([]);
  const [filterAddresses, setFilterAddresses] = useState<string[]>([]);
  const [filterGender, setFilterGender] = useState<'ALL' | 'M' | 'F'>('ALL');
  const [filterAddressType, setFilterAddressType] = useState<'ALL' | 'COMPANY' | 'OTHER'>('ALL');

  // Sorting States
  const [sortBy, setSortBy] = useState<
    'elapsed' | 'arrival' | 'passport' | 'name' | 'nationality' | 'address' | 'visa' | 'stayTo' | 'custom'
  >('elapsed');
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('desc');

  // Manual Selection & Custom Order state
  const [selectedPassports, setSelectedPassports] = useState<Record<string, boolean>>({});
  const [manualRowOrder, setManualRowOrder] = useState<string[]>([]);
  const [customRemarks, setCustomRemarks] = useState<Record<string, string>>({});
  const [isFilterPanelOpen, setIsFilterPanelOpen] = useState<boolean>(false);
  const [minElapsedDays, setMinElapsedDays] = useState<string>('');

  // Column Visibility state
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
  const [isPresetModalOpen, setIsPresetModalOpen] = useState<boolean>(false);

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
    showToast("မူလကော်လံပုံစံသို့ ပြန်ပြောင်းပြီးပါပြီ");
  };

  const handleApplyPreset = (preset: CRTPreset) => {
    setTableTitle(preset.tableTitle);
    setOfficerTitle(preset.officerTitle);
    setOfficerName(preset.officerName);
    if (preset.showSigner !== undefined) {
      setShowSigner(preset.showSigner);
      try { localStorage.setItem('crt_showSigner', String(preset.showSigner)); } catch {}
    }
    setUseBurmeseDigits(preset.useBurmeseDigits ?? true);
    setDataPool(preset.dataPool || 'STILL_IN');
    if (preset.visibleColumns && preset.visibleColumns.length > 0) {
      setVisibleColumns(preset.visibleColumns);
      try {
        localStorage.setItem('crt_visible_columns_v2', JSON.stringify(preset.visibleColumns));
      } catch {}
    }
    if (preset.filterNationalities) {
      setFilterNationalities(preset.filterNationalities);
    } else if (preset.filterNationality && preset.filterNationality !== 'ALL') {
      setFilterNationalities([preset.filterNationality]);
    } else {
      setFilterNationalities([]);
    }

    if (preset.filterVisaTypes) {
      setFilterVisaTypes(preset.filterVisaTypes);
    } else if (preset.filterVisaType && preset.filterVisaType !== 'ALL') {
      setFilterVisaTypes([preset.filterVisaType]);
    } else {
      setFilterVisaTypes([]);
    }

    if (preset.filterAddresses) {
      setFilterAddresses(preset.filterAddresses);
    } else if (preset.filterAddress && preset.filterAddress !== 'ALL') {
      setFilterAddresses([preset.filterAddress]);
    } else {
      setFilterAddresses([]);
    }

    // Restore Address Type (Other vs Co.Ltd vs All) & Gender filter
    if (preset.filterAddressType) {
      setFilterAddressType(preset.filterAddressType);
    } else {
      setFilterAddressType('ALL');
    }

    if (preset.filterGender) {
      setFilterGender(preset.filterGender);
    } else {
      setFilterGender('ALL');
    }

    if (preset.minElapsedDays !== undefined) {
      setMinElapsedDays(preset.minElapsedDays);
    } else {
      setMinElapsedDays('');
    }

    try {
      localStorage.setItem('crt_tableTitle', preset.tableTitle);
      localStorage.setItem('crt_officerTitle', preset.officerTitle);
      localStorage.setItem('crt_officerName', preset.officerName);
    } catch {}

    logActivity({
      action: 'UPDATE',
      module: 'CRT',
      details: `CRT Template Preset "${preset.name}" အား အသုံးပြုခဲ့သည်`
    });
  };

  const printRef = useRef<HTMLDivElement>(null);
  const previewContainerRef = useRef<HTMLDivElement>(null);
  const [containerWidth, setContainerWidth] = useState<number>(1122);
  const [viewZoomMode, setViewZoomMode] = useState<'fit' | 'full'>('fit');
  const [paperHeight, setPaperHeight] = useState<number>(800);

  useEffect(() => {
    if (!previewContainerRef.current) return;
    const updateDimensions = () => {
      if (previewContainerRef.current) {
        setContainerWidth(previewContainerRef.current.clientWidth);
      }
      if (printRef.current) {
        setPaperHeight(printRef.current.offsetHeight);
      }
    };
    updateDimensions();
    const observer = new ResizeObserver(updateDimensions);
    observer.observe(previewContainerRef.current);
    window.addEventListener('resize', updateDimensions);
    return () => {
      observer.disconnect();
      window.removeEventListener('resize', updateDimensions);
    };
  }, []);

  // Standard A4 landscape is 297mm ≈ 1122.5px
  const fitScale = useMemo(() => {
    if (containerWidth <= 0) return 1;
    const available = Math.max(280, containerWidth - 32);
    return Math.min(1, available / 1122.5);
  }, [containerWidth]);

  const currentScale = viewZoomMode === 'fit' ? fitScale : 1;

  // Helper: parse timestamp string to ms
  const parseTimestamp = (str?: string): number => {
    if (!str) return 0;
    try {
      const [datePart, timePart] = str.split(', ');
      if (!datePart) return 0;
      const [d, m, y] = datePart.split('/').map(Number);
      if (!d || !m || !y) return 0;
      let hh = 0, mm = 0, ss = 0;
      if (timePart) {
        const parts = timePart.split(':').map(Number);
        hh = parts[0] || 0;
        mm = parts[1] || 0;
        ss = parts[2] || 0;
      }
      return new Date(y, m - 1, d, hh, mm, ss).getTime();
    } catch {
      return 0;
    }
  };

  const isCoLtdAddress = (addressStr: string) => {
    const clean = (addressStr || '').toUpperCase();
    return clean.includes('CO') && (clean.includes('LTD') || clean.includes('L.T.D'));
  };

  const toBurmeseDigits = (numStr: string | number) => {
    const burmeseDigits = ['၀', '၁', '၂', '၃', '၄', '၅', '၆', '၇', '၈', '၉'];
    return String(numStr).replace(/[0-9]/g, (match) => burmeseDigits[parseInt(match)]);
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

  // Calculate elapsed days
  const computeElapsedDays = (arrTimeMs: number, refDateStr: string): number => {
    if (!arrTimeMs) return 0;
    let refMs = Date.now();
    if (refDateStr) {
      refMs = new Date(refDateStr + "T23:59:59").getTime();
    }
    const diff = refMs - arrTimeMs;
    if (diff < 0) return 1;
    return Math.floor(diff / (1000 * 60 * 60 * 24)) + 1;
  };

  // 1. Extract raw items based on data pool
  const rawList = useMemo(() => {
    const movements = Object.values(movementMap || {}) as MovementData[];
    
    if (dataPool === 'STILL_IN') {
      return movements.filter(m => !m.out);
    } else if (dataPool === 'INBOUND') {
      return movements.filter(m => !!m.in);
    } else if (dataPool === 'OUTBOUND') {
      return movements.filter(m => !!m.out);
    } else {
      return movements;
    }
  }, [movementMap, dataPool]);

  // Extract unique filter dropdown options
  const uniqueNationalities = useMemo(() => {
    const nats = new Set<string>();
    rawList.forEach(m => { if (m.nat) nats.add(m.nat.trim()); });
    return Array.from(nats).sort();
  }, [rawList]);

  const uniqueVisaTypes = useMemo(() => {
    const visas = new Set<string>();
    rawList.forEach(m => { if (m.visa) visas.add(m.visa.trim()); });
    return Array.from(visas).sort();
  }, [rawList]);

  const uniqueAddresses = useMemo(() => {
    const addrs = new Set<string>();
    rawList.forEach(m => { if (m.loc) addrs.add(m.loc.trim()); });
    return Array.from(addrs).sort();
  }, [rawList]);

  // 2. Filter list
  const filteredList = useMemo(() => {
    return rawList.filter(m => {
      const arrMs = m.inTime || parseTimestamp(m.in);

      // Date Range Filter
      if (dateRangeMode) {
        if (startDate) {
          const sMs = new Date(startDate + "T00:00:00").getTime();
          if (arrMs < sMs) return false;
        }
        if (endDate) {
          const eMs = new Date(endDate + "T23:59:59").getTime();
          if (arrMs > eMs) return false;
        }
      }

      // Nationality Filter (Multi-select)
      if (filterNationalities.length > 0) {
        const nat = (m.nat || '').trim();
        if (!filterNationalities.includes(nat)) return false;
      }

      // Visa Type Filter (Multi-select)
      if (filterVisaTypes.length > 0) {
        const visa = (m.visa || '').trim();
        if (!filterVisaTypes.includes(visa)) return false;
      }

      // Address Filter (Multi-select)
      if (filterAddresses.length > 0) {
        const loc = (m.loc || '').trim();
        if (!filterAddresses.includes(loc)) return false;
      }

      // Address Type Filter
      if (filterAddressType === 'COMPANY' && !isCoLtdAddress(m.loc)) {
        return false;
      }
      if (filterAddressType === 'OTHER' && isCoLtdAddress(m.loc)) {
        return false;
      }

      // Min Elapsed Days Filter
      if (minElapsedDays.trim()) {
        const minDays = parseInt(minElapsedDays.trim(), 10);
        if (!isNaN(minDays)) {
          const elapsed = computeElapsedDays(arrMs, selectedDate);
          if (elapsed < minDays) return false;
        }
      }

      // Gender Filter
      if (filterGender !== 'ALL' && m.gender !== filterGender) {
        return false;
      }

      // Live Text Search
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const match = 
          (m.p || '').toLowerCase().includes(q) ||
          (m.n || '').toLowerCase().includes(q) ||
          (m.nat || '').toLowerCase().includes(q) ||
          (m.loc || '').toLowerCase().includes(q) ||
          (m.visa || '').toLowerCase().includes(q) ||
          (m.vInfo || '').toLowerCase().includes(q);
        if (!match) return false;
      }

      return true;
    });
  }, [
    rawList, dateRangeMode, startDate, endDate,
    filterNationalities, filterVisaTypes, filterAddresses, 
    filterAddressType, filterGender, minElapsedDays, selectedDate, searchQuery
  ]);

  // 3. Transform to CustomReportRow items
  const processedRows = useMemo<CustomReportRow[]>(() => {
    return filteredList.map((m, idx) => {
      const arrMs = m.inTime || parseTimestamp(m.in);
      const elapsed = computeElapsedDays(arrMs, selectedDate);
      const isChecked = selectedPassports[m.p] !== undefined ? selectedPassports[m.p] : true;
      const userRemark = customRemarks[m.p] || '';

      return {
        id: m.p,
        passport: m.p,
        nationality: m.nat || '-',
        fullname: m.n || '-',
        gender: m.gender || '',
        visaType: m.visa || '-',
        visaNumber: m.visaNumber || '',
        dob: m.dob || '',
        stayFrom: m.start ? formatToDDMMYYYY(m.start) : '',
        stayTo: m.end ? formatToDDMMYYYY(m.end) : '',
        lastArrivalDate: m.in ? formatToDDMMYYYY(m.in) : '-',
        arrivalTimestampMs: arrMs,
        elapsedDays: elapsed,
        address: m.loc || '-',
        remarks: userRemark, // strictly no auto remark
        selected: isChecked,
        orderIndex: manualRowOrder.indexOf(m.p) !== -1 ? manualRowOrder.indexOf(m.p) : idx
      };
    });
  }, [filteredList, selectedDate, selectedPassports, customRemarks, manualRowOrder]);

  // 4. Sort rows
  const sortedRows = useMemo(() => {
    if (sortBy === 'custom') {
      return [...processedRows].sort((a, b) => a.orderIndex - b.orderIndex);
    }

    return [...processedRows].sort((a, b) => {
      let comparison = 0;
      switch (sortBy) {
        case 'elapsed':
          comparison = a.elapsedDays - b.elapsedDays;
          break;
        case 'arrival':
          comparison = a.arrivalTimestampMs - b.arrivalTimestampMs;
          break;
        case 'passport':
          comparison = a.passport.localeCompare(b.passport);
          break;
        case 'name':
          comparison = a.fullname.localeCompare(b.fullname);
          break;
        case 'nationality':
          comparison = a.nationality.localeCompare(b.nationality);
          break;
        case 'address':
          comparison = a.address.localeCompare(b.address);
          break;
        case 'visa':
          comparison = a.visaType.localeCompare(b.visaType);
          break;
        case 'stayTo':
          comparison = (a.stayTo || '').localeCompare(b.stayTo || '');
          break;
        default:
          comparison = 0;
      }
      return sortDirection === 'asc' ? comparison : -comparison;
    });
  }, [processedRows, sortBy, sortDirection]);

  // 5. Final active rows to show on the printed sheet (only selected ones)
  const displayRows = useMemo(() => {
    return sortedRows.filter(r => r.selected);
  }, [sortedRows]);

  // Selection handlers
  const toggleSelectAll = (select: boolean) => {
    const updated: Record<string, boolean> = {};
    sortedRows.forEach(r => {
      updated[r.passport] = select;
    });
    setSelectedPassports(prev => ({ ...prev, ...updated }));
    showToast(select ? `အားလုံး (${sortedRows.length}) ဦးကို ရွေးချယ်ပြီးပါပြီ` : "ရွေးချယ်မှု အားလုံးကို ပယ်ဖျက်ပြီးပါပြီ");
  };

  const toggleRowSelect = (passport: string) => {
    setSelectedPassports(prev => ({
      ...prev,
      [passport]: prev[passport] !== undefined ? !prev[passport] : false
    }));
  };

  // Header click sort handler
  const handleHeaderSort = (field: typeof sortBy) => {
    if (sortBy === field) {
      setSortDirection(prev => prev === 'asc' ? 'desc' : 'asc');
    } else {
      setSortBy(field);
      setSortDirection('desc');
    }
  };

  // Manual Row Move Up / Down
  const handleMoveRow = (passport: string, direction: 'up' | 'down') => {
    const currentList = sortedRows.map(r => r.passport);
    const index = currentList.indexOf(passport);
    if (index === -1) return;

    const targetIndex = direction === 'up' ? index - 1 : index + 1;
    if (targetIndex < 0 || targetIndex >= currentList.length) return;

    const temp = currentList[index];
    currentList[index] = currentList[targetIndex];
    currentList[targetIndex] = temp;

    setManualRowOrder(currentList);
    setSortBy('custom');
    showToast(`အစီအစဉ် ပြောင်းလဲပြီးပါပြီ`);
  };

  // Save Title
  const handleSaveTitle = (newTitle: string) => {
    setTableTitle(newTitle);
    localStorage.setItem('crt_tableTitle', newTitle);
    setIsEditingTitle(false);
    showToast("ဇယားခေါင်းစဉ် သိမ်းဆည်းပြီးပါပြီ");
  };

  // Save Officer Title and Name
  const handleSaveOfficerTitle = (newTitle: string) => {
    setOfficerTitle(newTitle);
    localStorage.setItem('crt_officerTitle', newTitle);
  };

  const handleSaveOfficerName = (newName: string) => {
    setOfficerName(newName);
    localStorage.setItem('crt_officerName', newName);
  };

  // Handle Remark change per row
  const handleRemarkChange = (passport: string, text: string) => {
    setCustomRemarks(prev => ({ ...prev, [passport]: text }));
  };

  // Export to Excel
  const exportToExcel = () => {
    if (displayRows.length === 0) {
      showToast("ထုတ်ယူရန် အချက်အလက်မရှိပါ");
      return;
    }

    const activeColDefs = CRT_COLUMNS_DEFINITION.filter(col => visibleColumns.includes(col.id));

    const excelData = displayRows.map((r, idx) => {
      const stayPeriod = (r.stayFrom || r.stayTo) ? `${r.stayFrom || '-'} မှ ${r.stayTo || '-'}` : '-';
      const genderStr = r.gender === 'M' ? 'ကျား' : r.gender === 'F' ? 'မ' : '-';
      const rowObj: Record<string, any> = {};

      activeColDefs.forEach(col => {
        switch (col.id) {
          case 'sr':
            rowObj[col.label] = idx + 1;
            break;
          case 'passport':
            rowObj[col.label] = r.passport;
            break;
          case 'nationality':
            rowObj[col.label] = r.nationality;
            break;
          case 'fullname':
            rowObj[col.label] = r.fullname;
            break;
          case 'gender':
            rowObj[col.label] = genderStr;
            break;
          case 'visaType':
            rowObj[col.label] = r.visaType;
            break;
          case 'stayPeriod':
            rowObj[col.label] = stayPeriod;
            break;
          case 'arrivalDate':
            rowObj[col.label] = r.lastArrivalDate;
            break;
          case 'elapsedDays':
            rowObj[col.label] = r.elapsedDays;
            break;
          case 'address':
            rowObj[col.label] = r.address;
            break;
          case 'remarks':
            rowObj[col.label] = r.remarks || '';
            break;
        }
      });

      return rowObj;
    });

    const worksheet = XLSX.utils.json_to_sheet(excelData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Custom_Report");
    
    // Auto-fit column width
    worksheet['!cols'] = activeColDefs.map(c => ({ wch: c.excelWidth }));

    const safeFileName = tableTitle.replace(/[\s\/\\]+/g, '_').substring(0, 30);
    XLSX.writeFile(workbook, `${safeFileName}_${selectedDate}.xlsx`);

    logActivity({
      action: 'UPDATE',
      module: 'CRT',
      details: `CRT Excel ထုတ်ယူခဲ့သည် (${displayRows.length} ဦး, ${activeColDefs.length} ကော်လံ)`
    });

    showToast("Excel ဖိုင် ထုတ်ယူပြီးပါပြီ");
  };

  // Export to Photo (PNG)
  const exportToPhoto = async () => {
    const element = printRef.current;
    if (!element) {
      showToast("ဇယားမတွေ့ရှိပါ။");
      return;
    }

    try {
      showToast("ပုံ (PNG) အဖြစ် သိမ်းဆည်းနေပါသည်...");
      // Dynamic import html2canvas
      const html2canvasModule = await import('html2canvas');
      const html2canvas = html2canvasModule.default;

      const canvas = await html2canvas(element, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#ffffff',
        logging: false
      });

      const dataUrl = canvas.toDataURL('image/png');
      const link = document.createElement('a');
      const safeFileName = tableTitle.replace(/[\s\/\\]+/g, '_').substring(0, 30);
      link.download = `${safeFileName}_${selectedDate}.png`;
      link.href = dataUrl;
      link.click();
      showToast("ပုံ (PNG) အဖြစ် သိမ်းဆည်းပြီးပါပြီ");
    } catch (err) {
      console.error('Export photo error:', err);
      showToast("ပုံထုတ်ယူရာတွင် ချို့ယွင်းချက်ရှိပါသည်");
    }
  };

  // Direct Print
  const triggerPrint = () => {
    const element = printRef.current;
    if (!element) {
      showToast("Print target not found");
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
    styleEl.innerHTML = `
      @media print { 
        @page { size: A4 landscape; margin: 0.4in; }
        body { font-family: 'Pyidaungsu', 'Pyidaungsu Number', sans-serif !important; }
        table { page-break-inside: auto; }
        tr { page-break-inside: avoid; page-break-after: auto; }
        thead { display: table-header-group; }
        tfoot { display: table-footer-group; }
        input, textarea { border: none !important; background: transparent !important; box-shadow: none !important; outline: none !important; }
      }
    `;
    const clone = element.cloneNode(true) as HTMLElement;
    const originalInputs = element.querySelectorAll('input, textarea, select');
    const clonedInputs = clone.querySelectorAll('input, textarea, select');
    originalInputs.forEach((orig: any, i) => {
      if (clonedInputs[i]) {
        (clonedInputs[i] as HTMLInputElement).value = orig.value;
      }
    });
    clone.style.display = 'block';
    clone.style.visibility = 'visible';
    printArea.appendChild(styleEl);
    printArea.appendChild(clone);
    window.print();
  };

  return (
    <div className="w-full max-w-7xl mx-auto py-4 sm:py-6 px-2 sm:px-4 space-y-4 sm:space-y-6 overflow-hidden">
      {/* Top Banner & Control Deck */}
      <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 rounded-3xl p-6 text-white shadow-xl border border-indigo-900/50">
        <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-6">
          <div className="space-y-1.5">
            <div className="flex items-center gap-3">
              <div className="p-2.5 bg-indigo-500/20 text-indigo-400 rounded-2xl border border-indigo-500/30">
                <FileSpreadsheet size={24} />
              </div>
              <div>
                <h1 className="text-xl md:text-2xl font-black tracking-tight text-white flex items-center gap-2">
                  <span>စိတ်ကြိုက် အစီရင်ခံစာ ဇယားသစ် (Custom Report Table)</span>
                  <span className="text-[10px] uppercase font-black px-2.5 py-0.5 rounded-full bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                    A4 Landscape
                  </span>
                </h1>
                <p className="text-xs text-indigo-200/80 font-medium">
                  OTHER ADDRESS RESIDENTS စာတိုင်ပုံစံအတိုင်း ခေါင်းစဉ်စိတ်ကြိုက်ပြင်ဆင်နိုင်ပြီး Manual Sorting & Filtering အပြည့်အဝ ပြုလုပ်နိုင်ပါသည်
                </p>
              </div>
            </div>
          </div>

          {/* Action Buttons */}
          <div className="flex flex-wrap items-center gap-2.5">
            <button
              onClick={() => setIsPresetModalOpen(true)}
              className="px-3.5 py-2.5 text-xs font-black rounded-xl bg-purple-600 hover:bg-purple-500 text-white shadow-lg transition-all flex items-center gap-1.5 cursor-pointer"
              title="Report Template Presets (အစီရင်ခံစာ ပုံစံများ ရွေးချယ်/သိမ်းမည်)"
            >
              <Bookmark size={15} />
              <span>Templates</span>
            </button>

            <button
              onClick={() => setIsColumnModalOpen(true)}
              className="px-3.5 py-2.5 text-xs font-black rounded-xl bg-slate-800 hover:bg-slate-700 text-indigo-200 border border-indigo-500/40 shadow-lg transition-all flex items-center gap-1.5 cursor-pointer"
              title="Column Visibility (ကော်လံများ စိတ်ကြိုက် ဖွင့်/ပိတ်ခြင်း)"
            >
              <Columns size={15} />
              <span>Columns ({visibleColumns.length})</span>
            </button>

            <button
              onClick={exportToExcel}
              className="px-4 py-2.5 text-xs font-black rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white shadow-lg transition-all flex items-center gap-2 cursor-pointer"
              title="Export to Excel (.xlsx)"
            >
              <FileSpreadsheet size={15} />
              <span>Export Excel</span>
            </button>

            <button
              onClick={exportToPhoto}
              className="px-4 py-2.5 text-xs font-black rounded-xl bg-teal-600 hover:bg-teal-500 text-white shadow-lg transition-all flex items-center gap-2 cursor-pointer"
              title="Save as High-Res PNG Photo"
            >
              <Download size={15} />
              <span>Save as Photo</span>
            </button>

            <button
              onClick={triggerPrint}
              className="px-4 py-2.5 text-xs font-black rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white shadow-lg transition-all flex items-center gap-2 cursor-pointer"
              title="Print Table (A4 Landscape)"
            >
              <Printer size={15} />
              <span>Print Table (Landscape)</span>
            </button>
          </div>
        </div>

        {/* Quick Title Editor & Config Strip */}
        <div className="mt-6 pt-5 border-t border-white/10 grid grid-cols-1 md:grid-cols-3 gap-4 items-center">
          {/* Title Editor */}
          <div className="md:col-span-2 flex items-center gap-3 bg-white/5 p-2.5 rounded-2xl border border-white/10">
            <Edit3 size={18} className="text-indigo-400 shrink-0 ml-2" />
            <div className="flex-1">
              <label className="block text-[10px] font-black uppercase text-indigo-300 tracking-wider">
                ဇယားခေါင်းစဉ် (Editable Table Title)
              </label>
              {isEditingTitle ? (
                <div className="flex items-center gap-2 mt-1">
                  <input
                    type="text"
                    defaultValue={tableTitle}
                    id="crt_title_input"
                    className="flex-1 bg-white text-slate-900 px-3 py-1.5 rounded-xl font-bold text-sm outline-none font-pyidaungsu"
                    autoFocus
                  />
                  <button
                    onClick={() => {
                      const input = document.getElementById('crt_title_input') as HTMLInputElement;
                      if (input) handleSaveTitle(input.value);
                    }}
                    className="p-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl cursor-pointer"
                    title="သိမ်းဆည်းမည်"
                  >
                    <Check size={16} />
                  </button>
                </div>
              ) : (
                <div 
                  onClick={() => setIsEditingTitle(true)}
                  className="font-bold text-sm text-white hover:text-indigo-200 cursor-pointer flex items-center gap-2 group mt-0.5 font-pyidaungsu"
                  title="ခေါင်းစဉ်ပြင်ရန် နှိပ်ပါ"
                >
                  <span>{tableTitle}</span>
                  <span className="text-[10px] text-indigo-300 opacity-70 group-hover:opacity-100 font-normal">
                    (Click to edit)
                  </span>
                </div>
              )}
            </div>
          </div>

          {/* Reference Date & Burmese Digits Toggle */}
          <div className="flex items-center justify-between md:justify-end gap-3">
            <div>
              <label className="block text-[10px] font-black uppercase text-indigo-300 tracking-wider mb-1">
                ရည်ညွှန်းရက်စွဲ (Date)
              </label>
              <input
                type="date"
                value={selectedDate}
                onChange={(e) => setSelectedDate(e.target.value)}
                className="bg-white text-slate-900 px-3 py-1.5 rounded-xl font-bold text-xs outline-none"
              />
            </div>

            <button
              onClick={() => setUseBurmeseDigits(!useBurmeseDigits)}
              className={`px-3 py-2 text-xs font-black rounded-xl border transition-all flex items-center gap-1.5 cursor-pointer mt-4 ${
                useBurmeseDigits 
                  ? 'bg-amber-500/20 text-amber-300 border-amber-500/40' 
                  : 'bg-white/10 text-white border-white/20'
              }`}
            >
              <span>မြန်မာဂဏန်း:</span>
              <span className="font-bold">{useBurmeseDigits ? 'ON (၁, ၂, ၃)' : 'OFF (1, 2, 3)'}</span>
            </button>
          </div>
        </div>

        {/* Officer Signature Customizer Strip */}
        <div className="mt-4 pt-4 border-t border-white/10 flex flex-col md:flex-row gap-4 items-start md:items-center justify-between">
          <div className="flex items-center gap-3">
            <button
              type="button"
              onClick={() => handleToggleShowSigner(!showSigner)}
              className={`px-3.5 py-2 rounded-xl font-black text-xs transition-all flex items-center gap-2 cursor-pointer border shadow-sm ${
                showSigner
                  ? 'bg-emerald-500/25 text-emerald-200 border-emerald-400/40 hover:bg-emerald-500/35'
                  : 'bg-rose-500/20 text-rose-300 border-rose-500/40 hover:bg-rose-500/30'
              }`}
            >
              <Shield size={14} className={showSigner ? "text-emerald-400" : "text-rose-400"} />
              <span>{showSigner ? '✓ Signer လက်မှတ်ပါဝင်မည်' : '✕ Signer ဖြုတ်ထားပါသည် (Hidden)'}</span>
            </button>
            <span className="text-[11px] text-indigo-300 font-bold hidden sm:inline">
              {showSigner ? 'အစီရင်ခံစာအောက်ခြေတွင် အရာရှိလက်မှတ် ထည့်သွင်းပြသမည်' : 'အစီရင်ခံစာအောက်ခြေတွင် အရာရှိလက်မှတ် မပြပါ'}
            </span>
          </div>

          {showSigner && (
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 w-full md:w-auto md:min-w-[420px]">
              <div>
                <label className="block text-[10px] font-black uppercase text-indigo-300 tracking-wider mb-1">
                  ရာထူး (Officer Title)
                </label>
                <input
                  type="text"
                  value={officerTitle}
                  onChange={(e) => handleSaveOfficerTitle(e.target.value)}
                  list="crtOfficialTitlesList"
                  placeholder="ဥပမာ- လဝကမှူး..."
                  className="w-full bg-white text-slate-900 px-3 py-1.5 rounded-xl font-bold text-xs outline-none font-pyidaungsu border border-indigo-200 shadow-inner"
                />
              </div>
              <div>
                <label className="block text-[10px] font-black uppercase text-indigo-300 tracking-wider mb-1">
                  အမည် (Officer Name)
                </label>
                <input
                  type="text"
                  value={officerName}
                  onChange={(e) => handleSaveOfficerName(e.target.value)}
                  list="crtOfficialNamesList"
                  placeholder="ဥပမာ- ဒုတိယလဝကမှူး / အမည်..."
                  className="w-full bg-white text-slate-900 px-3 py-1.5 rounded-xl font-bold text-xs outline-none font-pyidaungsu border border-indigo-200 shadow-inner"
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

      {/* Interactive Controls & Filters Card */}
      <div className="bg-white rounded-3xl p-5 shadow-md border border-slate-200/80 space-y-4">
        <div className="flex flex-wrap items-center justify-between gap-3 pb-3 border-b border-slate-100">
          {/* Data Pool Selector */}
          <div className="flex items-center gap-1.5 bg-slate-100 p-1 rounded-2xl">
            <button
              onClick={() => setDataPool('STILL_IN')}
              className={`px-3.5 py-1.5 rounded-xl font-black text-xs transition-all cursor-pointer ${
                dataPool === 'STILL_IN'
                  ? 'bg-indigo-600 text-white shadow-md'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              လက်ရှိနေထိုင်ဆဲ (Still In)
            </button>
            <button
              onClick={() => setDataPool('ALL_MOVEMENTS')}
              className={`px-3.5 py-1.5 rounded-xl font-black text-xs transition-all cursor-pointer ${
                dataPool === 'ALL_MOVEMENTS'
                  ? 'bg-slate-800 text-white shadow-md'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              အားလုံး (All Records)
            </button>
            <button
              onClick={() => setDataPool('INBOUND')}
              className={`px-3.5 py-1.5 rounded-xl font-black text-xs transition-all cursor-pointer ${
                dataPool === 'INBOUND'
                  ? 'bg-emerald-700 text-white shadow-md'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              ဝင်ရောက်သူများ (Inbound)
            </button>
            <button
              onClick={() => setDataPool('OUTBOUND')}
              className={`px-3.5 py-1.5 rounded-xl font-black text-xs transition-all cursor-pointer ${
                dataPool === 'OUTBOUND'
                  ? 'bg-orange-600 text-white shadow-md'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              ထွက်ခွါသူများ (Outbound)
            </button>
          </div>

          {/* Quick Selection Toolbar */}
          <div className="flex items-center gap-2">
            <span className="text-xs font-bold text-slate-500">
              ဇယားတွင် ပါဝင်သူ: <strong className="text-indigo-900 font-black">{displayRows.length}</strong> / {sortedRows.length} ဦး
            </span>
            <button
              onClick={() => toggleSelectAll(true)}
              className="px-3 py-1.5 text-xs font-bold rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 flex items-center gap-1 cursor-pointer"
            >
              <CheckSquare size={13} />
              <span>အားလုံးရွေးမည်</span>
            </button>
            <button
              onClick={() => toggleSelectAll(false)}
              className="px-3 py-1.5 text-xs font-bold rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 flex items-center gap-1 cursor-pointer"
            >
              <Square size={13} />
              <span>အားလုံးပယ်ဖျက်</span>
            </button>
            <button
              onClick={() => setIsFilterPanelOpen(!isFilterPanelOpen)}
              className={`px-3.5 py-1.5 text-xs font-bold rounded-xl flex items-center gap-1.5 cursor-pointer border transition-all ${
                isFilterPanelOpen 
                  ? 'bg-indigo-50 border-indigo-300 text-indigo-700 font-black' 
                  : 'bg-white border-slate-300 text-slate-700 hover:bg-slate-50'
              }`}
            >
              <Filter size={14} />
              <span>Manual Filters</span>
              <ChevronDown size={14} className={`transition-transform ${isFilterPanelOpen ? 'rotate-180' : ''}`} />
            </button>
          </div>
        </div>

        {/* Quick Search & Sort Bar */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
          {/* Live Search */}
          <div className="relative">
            <Search size={16} className="absolute left-3.5 top-3 text-slate-400" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="ရှာဖွေရန် (ပတ်စပို့၊ အမည်၊ လိပ်စာ၊ ဗီဇာ၊ နိုင်ငံသား)..."
              className="w-full pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:bg-white focus:border-indigo-500 outline-none transition-all"
            />
          </div>

          {/* Manual Sort Field */}
          <div className="flex items-center gap-2">
            <div className="flex-1 relative">
              <label className="text-[10px] font-black text-slate-400 uppercase absolute left-3 top-1">
                Manual Sorting
              </label>
              <select
                value={sortBy}
                onChange={(e) => setSortBy(e.target.value as any)}
                className="w-full pt-4 pb-1.5 px-3 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-indigo-900 outline-none cursor-pointer"
              >
                <option value="elapsed">ရက်ပေါင်းများသူမှ နည်းသူသို့ (Elapsed Days)</option>
                <option value="arrival">ရောက်ရှိရက်စွဲ (Arrival Date)</option>
                <option value="passport">ပတ်စပို့နံပါတ် (Passport No)</option>
                <option value="name">အမည် (Full Name)</option>
                <option value="nationality">နိုင်ငံအမည် (Nationality)</option>
                <option value="address">တည်းခိုလိပ်စာ (Stay Address)</option>
                <option value="visa">ဗီဇာအမျိုးအစား (Visa Type)</option>
                <option value="stayTo">ဗီဇာသက်တမ်းကုန်ရက် (Stay Expiry)</option>
                <option value="custom">စိတ်ကြိုက်ရွှေ့ထားသည့် အစီအစဉ် (Custom Order)</option>
              </select>
            </div>

            <button
              onClick={() => setSortDirection(prev => prev === 'asc' ? 'desc' : 'asc')}
              className="p-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl cursor-pointer border border-slate-200"
              title={sortDirection === 'asc' ? 'ငယ်စဉ်ကြီးလိုက် (Ascending)' : 'ကြီးစဉ်ငယ်လိုက် (Descending)'}
            >
              <ArrowUpDown size={16} />
            </button>
          </div>

          {/* Reset Filters */}
          <div className="flex items-center gap-2">
            <button
              onClick={() => {
                setSearchQuery('');
                setFilterNationalities([]);
                setFilterVisaTypes([]);
                setFilterAddresses([]);
                setFilterGender('ALL');
                setFilterAddressType('ALL');
                setDateRangeMode(false);
                setStartDate('');
                setEndDate('');
                setMinElapsedDays('');
                setSortBy('elapsed');
                setSortDirection('desc');
                showToast("Filters Reset ပြုလုပ်ပြီးပါပြီ");
              }}
              className="flex-1 py-2 px-3 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-xl font-bold text-xs flex items-center justify-center gap-1.5 cursor-pointer transition-all"
            >
              <RotateCcw size={14} />
              <span>Reset All Filters</span>
            </button>
          </div>
        </div>

        {/* Expanded Manual Filter Panel (Multi-Selectable) */}
        {isFilterPanelOpen && (
          <div className="p-4 bg-slate-50 border border-indigo-100 rounded-2xl space-y-4 animate-in fade-in duration-200">
            <div className="flex items-center justify-between pb-2 border-b border-slate-200/80">
              <span className="text-xs font-black text-indigo-950 uppercase tracking-wider flex items-center gap-1.5">
                <Filter size={14} className="text-indigo-600" />
                <span>Multiple Selectable Filters (အများအပြား ရွေးချယ်စစ်ထုတ်ရန်)</span>
              </span>
              {(filterNationalities.length > 0 || filterVisaTypes.length > 0 || filterAddresses.length > 0) && (
                <button
                  type="button"
                  onClick={() => {
                    setFilterNationalities([]);
                    setFilterVisaTypes([]);
                    setFilterAddresses([]);
                    showToast("Multi-select Filters များ ရှင်းလင်းပြီးပါပြီ");
                  }}
                  className="text-[11px] font-bold text-rose-600 hover:text-rose-700 underline cursor-pointer"
                >
                  Clear Selection
                </button>
              )}
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              {/* Nationality Multi-Select */}
              <div className="bg-white p-3 rounded-xl border border-slate-200 flex flex-col">
                <div className="flex items-center justify-between mb-2">
                  <label className="text-[11px] font-black text-slate-700 uppercase">
                    နိုင်ငံအမည် (Nationality) {filterNationalities.length > 0 && <span className="text-indigo-600 font-extrabold">({filterNationalities.length} ရွေးထား)</span>}
                  </label>
                  <div className="flex items-center gap-1 text-[10px]">
                    <button
                      type="button"
                      onClick={() => setFilterNationalities([...uniqueNationalities])}
                      className="text-indigo-600 hover:underline font-bold"
                    >
                      All
                    </button>
                    <span className="text-slate-300">|</span>
                    <button
                      type="button"
                      onClick={() => setFilterNationalities([])}
                      className="text-slate-500 hover:underline font-bold"
                    >
                      None
                    </button>
                  </div>
                </div>
                <div className="max-h-36 overflow-y-auto space-y-1 pr-1 border border-slate-100 rounded-lg p-1.5 bg-slate-50/50">
                  {uniqueNationalities.length === 0 ? (
                    <div className="text-[11px] text-slate-400 p-2 text-center italic">နိုင်ငံအမည် မရှိပါ</div>
                  ) : (
                    uniqueNationalities.map(n => {
                      const isSelected = filterNationalities.includes(n);
                      return (
                        <label
                          key={n}
                          className={`flex items-center justify-between px-2 py-1 rounded-md text-xs font-bold cursor-pointer transition-colors ${
                            isSelected ? 'bg-indigo-600 text-white shadow-xs' : 'hover:bg-slate-200 text-slate-700'
                          }`}
                        >
                          <span>{n}</span>
                          <input
                            type="checkbox"
                            checked={isSelected}
                            onChange={() => {
                              setFilterNationalities(prev => 
                                prev.includes(n) ? prev.filter(x => x !== n) : [...prev, n]
                              );
                            }}
                            className="rounded text-indigo-600 shrink-0 ml-2"
                          />
                        </label>
                      );
                    })
                  )}
                </div>
              </div>

              {/* Visa Type Multi-Select */}
              <div className="bg-white p-3 rounded-xl border border-slate-200 flex flex-col">
                <div className="flex items-center justify-between mb-2">
                  <label className="text-[11px] font-black text-slate-700 uppercase">
                    ဗီဇာအမျိုးအစား (Visa Type) {filterVisaTypes.length > 0 && <span className="text-indigo-600 font-extrabold">({filterVisaTypes.length} ရွေးထား)</span>}
                  </label>
                  <div className="flex items-center gap-1 text-[10px]">
                    <button
                      type="button"
                      onClick={() => setFilterVisaTypes([...uniqueVisaTypes])}
                      className="text-indigo-600 hover:underline font-bold"
                    >
                      All
                    </button>
                    <span className="text-slate-300">|</span>
                    <button
                      type="button"
                      onClick={() => setFilterVisaTypes([])}
                      className="text-slate-500 hover:underline font-bold"
                    >
                      None
                    </button>
                  </div>
                </div>
                <div className="max-h-36 overflow-y-auto space-y-1 pr-1 border border-slate-100 rounded-lg p-1.5 bg-slate-50/50">
                  {uniqueVisaTypes.length === 0 ? (
                    <div className="text-[11px] text-slate-400 p-2 text-center italic">ဗီဇာအမျိုးအစား မရှိပါ</div>
                  ) : (
                    uniqueVisaTypes.map(v => {
                      const isSelected = filterVisaTypes.includes(v);
                      return (
                        <label
                          key={v}
                          className={`flex items-center justify-between px-2 py-1 rounded-md text-xs font-bold cursor-pointer transition-colors ${
                            isSelected ? 'bg-indigo-600 text-white shadow-xs' : 'hover:bg-slate-200 text-slate-700'
                          }`}
                        >
                          <span>{v}</span>
                          <input
                            type="checkbox"
                            checked={isSelected}
                            onChange={() => {
                              setFilterVisaTypes(prev => 
                                prev.includes(v) ? prev.filter(x => x !== v) : [...prev, v]
                              );
                            }}
                            className="rounded text-indigo-600 shrink-0 ml-2"
                          />
                        </label>
                      );
                    })
                  )}
                </div>
              </div>

              {/* Address / Location Multi-Select */}
              <div className="bg-white p-3 rounded-xl border border-slate-200 flex flex-col">
                <div className="flex items-center justify-between mb-2">
                  <label className="text-[11px] font-black text-slate-700 uppercase">
                    တည်းခိုလိပ်စာ (Address / Hotel) {filterAddresses.length > 0 && <span className="text-indigo-600 font-extrabold">({filterAddresses.length} ရွေးထား)</span>}
                  </label>
                  <div className="flex items-center gap-1 text-[10px]">
                    <button
                      type="button"
                      onClick={() => setFilterAddresses([...uniqueAddresses])}
                      className="text-indigo-600 hover:underline font-bold"
                    >
                      All
                    </button>
                    <span className="text-slate-300">|</span>
                    <button
                      type="button"
                      onClick={() => setFilterAddresses([])}
                      className="text-slate-500 hover:underline font-bold"
                    >
                      None
                    </button>
                  </div>
                </div>
                <div className="max-h-36 overflow-y-auto space-y-1 pr-1 border border-slate-100 rounded-lg p-1.5 bg-slate-50/50">
                  {uniqueAddresses.length === 0 ? (
                    <div className="text-[11px] text-slate-400 p-2 text-center italic">လိပ်စာမရှိပါ</div>
                  ) : (
                    uniqueAddresses.map(addr => {
                      const isSelected = filterAddresses.includes(addr);
                      return (
                        <label
                          key={addr}
                          className={`flex items-center justify-between px-2 py-1 rounded-md text-xs font-bold cursor-pointer transition-colors ${
                            isSelected ? 'bg-indigo-600 text-white shadow-xs' : 'hover:bg-slate-200 text-slate-700'
                          }`}
                        >
                          <span className="truncate max-w-[200px]" title={addr}>{addr}</span>
                          <input
                            type="checkbox"
                            checked={isSelected}
                            onChange={() => {
                              setFilterAddresses(prev => 
                                prev.includes(addr) ? prev.filter(x => x !== addr) : [...prev, addr]
                              );
                            }}
                            className="rounded text-indigo-600 shrink-0 ml-2"
                          />
                        </label>
                      );
                    })
                  )}
                </div>
              </div>
            </div>

            {/* Other Secondary Filters */}
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-2 border-t border-slate-200/80 items-end">
              {/* Address Category */}
              <div>
                <label className="block text-[11px] font-black text-slate-700 uppercase mb-1">
                  လိပ်စာအမျိုးအစား (Address Category)
                </label>
                <select
                  value={filterAddressType}
                  onChange={(e) => setFilterAddressType(e.target.value as any)}
                  className="w-full p-2 bg-white border border-slate-200 rounded-xl text-xs font-bold text-slate-800 outline-none"
                >
                  <option value="ALL">လိပ်စာအားလုံး (All Locations)</option>
                  <option value="OTHER">အခြားလိပ်စာ/ဟိုတယ် (Other / Hotel)</option>
                  <option value="COMPANY">ကုမ္ပဏီ (Company / Co.Ltd)</option>
                </select>
              </div>

              {/* Gender Filter */}
              <div>
                <label className="block text-[11px] font-black text-slate-700 uppercase mb-1">
                  ကျား / မ (Gender)
                </label>
                <select
                  value={filterGender}
                  onChange={(e) => setFilterGender(e.target.value as any)}
                  className="w-full p-2 bg-white border border-slate-200 rounded-xl text-xs font-bold text-slate-800 outline-none"
                >
                  <option value="ALL">ကျား/မ အားလုံး (All)</option>
                  <option value="M">ကျား (Male)</option>
                  <option value="F">မ (Female)</option>
                </select>
              </div>

              {/* Min Elapsed Days */}
              <div>
                <label className="block text-[11px] font-black text-slate-700 uppercase mb-1">
                  အနည်းဆုံးနေထိုင်ရက် (Min Days)
                </label>
                <input
                  type="number"
                  value={minElapsedDays}
                  onChange={(e) => setMinElapsedDays(e.target.value)}
                  placeholder="ဥပမာ- 30"
                  className="w-full p-2 bg-white border border-slate-200 rounded-xl text-xs font-bold text-slate-800 outline-none"
                />
              </div>
            </div>

            {/* Date Range Mode */}
            <div className="flex flex-wrap items-center gap-3 pt-2 border-t border-slate-200/60">
              <label className="flex items-center gap-2 cursor-pointer text-xs font-black text-indigo-900">
                <input
                  type="checkbox"
                  checked={dateRangeMode}
                  onChange={(e) => setDateRangeMode(e.target.checked)}
                  className="rounded text-indigo-600"
                />
                <span>ရက်စွဲအပိုင်းအခြားဖြင့် စစ်ထုတ်ရန် (Date Range Mode)</span>
              </label>

              {dateRangeMode && (
                <div className="flex items-center gap-2">
                  <input
                    type="date"
                    value={startDate}
                    onChange={(e) => setStartDate(e.target.value)}
                    className="p-1.5 bg-white border border-slate-300 rounded-lg text-xs font-bold"
                    placeholder="Start Date"
                  />
                  <span className="text-slate-400 font-bold">မှ</span>
                  <input
                    type="date"
                    value={endDate}
                    onChange={(e) => setEndDate(e.target.value)}
                    className="p-1.5 bg-white border border-slate-300 rounded-lg text-xs font-bold"
                    placeholder="End Date"
                  />
                </div>
              )}
            </div>
          </div>
        )}
      </div>

      {/* Row Selection & Manual Ordering Drawer */}
      <div className="bg-slate-50 border border-slate-200 rounded-3xl p-4">
        <div className="flex items-center justify-between mb-3">
          <div className="flex items-center gap-2">
            <span className="text-xs font-black uppercase text-slate-700 tracking-wider">
              📋 Manual Row Selection & Order Manager ({sortedRows.length} items)
            </span>
            <span className="text-[10px] bg-indigo-100 text-indigo-800 font-bold px-2 py-0.5 rounded-full">
              ဇယားတွင် ပါဝင်စေလိုသူများကို Checkbox ဖြင့် ရွေးချယ်နိုင်ပြီး ▲ / ▼ ဖြင့် အစီအစဉ်ရွှေ့နိုင်ပါသည်
            </span>
          </div>
        </div>

        <div className="max-h-56 overflow-y-auto bg-white rounded-2xl border border-slate-200 divide-y divide-slate-100">
          {sortedRows.length === 0 ? (
            <div className="p-4 text-center text-xs text-slate-400 italic">
              စစ်ထုတ်ထားသော အချက်အလက် မရှိပါ။
            </div>
          ) : (
            sortedRows.map((r, idx) => (
              <div 
                key={r.passport} 
                className={`p-2.5 flex items-center justify-between text-xs transition-colors ${
                  r.selected ? 'bg-indigo-50/30' : 'bg-slate-50/50 opacity-60'
                }`}
              >
                <div className="flex items-center gap-3 flex-1 min-w-0">
                  <input
                    type="checkbox"
                    checked={r.selected}
                    onChange={() => toggleRowSelect(r.passport)}
                    className="w-4 h-4 rounded text-indigo-600 cursor-pointer"
                  />
                  <span className="font-bold text-slate-400 w-6 text-center">{idx + 1}</span>
                  <span className="font-mono font-black text-indigo-950">{r.passport}</span>
                  <span className="font-bold text-slate-800 truncate max-w-[180px]">{r.fullname}</span>
                  <span className="text-[10px] bg-slate-100 text-slate-600 px-1.5 py-0.5 rounded font-bold">{r.nationality}</span>
                  <span className="text-[10px] text-slate-500 truncate max-w-[200px]">{r.address}</span>
                  <span className="text-[10px] font-bold text-indigo-700">{r.elapsedDays} ရက်</span>
                </div>

                <div className="flex items-center gap-1">
                  <button
                    onClick={() => handleMoveRow(r.passport, 'up')}
                    disabled={idx === 0}
                    className="p-1 rounded bg-slate-100 hover:bg-slate-200 text-slate-600 disabled:opacity-30 cursor-pointer"
                    title="အပေါ်သို့ ရွှေ့မည်"
                  >
                    <ArrowUp size={13} />
                  </button>
                  <button
                    onClick={() => handleMoveRow(r.passport, 'down')}
                    disabled={idx === sortedRows.length - 1}
                    className="p-1 rounded bg-slate-100 hover:bg-slate-200 text-slate-600 disabled:opacity-30 cursor-pointer"
                    title="အောက်သို့ ရွှေ့မည်"
                  >
                    <ArrowDown size={13} />
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
        <div className="w-full max-w-[297mm] flex flex-wrap justify-between items-center gap-3 mb-4">
          <div className="flex items-center gap-2 flex-wrap">
            <span className="text-xs font-black uppercase text-slate-700 tracking-wider">
              Custom Report Sheet (A4 Landscape Preview)
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
                title="မျက်နှာပြင်နှင့် အချိုးကျ အပြည့်ချုံ့ကြည့်မည် (Fit Screen)"
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
                title="မူလအရွယ် 100% ဖြင့် ဘယ်ညာ ဆွဲဖတ်မည် (Full 100% Scroll)"
              >
                <Maximize2 size={12} />
                <span>100% Full Size</span>
              </button>
            </div>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <button
              onClick={exportToExcel}
              className="px-3.5 py-2 text-xs font-bold rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer"
            >
              <FileSpreadsheet size={13} />
              <span>Export Excel</span>
            </button>
            <button
              onClick={exportToPhoto}
              className="px-3.5 py-2 text-xs font-bold rounded-xl bg-teal-600 hover:bg-teal-700 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer"
            >
              <Download size={13} />
              <span>Save as Photo</span>
            </button>
            <button
              onClick={triggerPrint}
              className="px-3.5 py-2 text-xs font-bold rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white shadow-md transition-all flex items-center gap-1.5 cursor-pointer"
            >
              <Printer size={13} />
              <span>Print Table</span>
            </button>
          </div>
        </div>

        {/* Mobile Swipe Hint when in 100% Full View */}
        {viewZoomMode === 'full' && fitScale < 1 && (
          <div className="w-full max-w-[297mm] text-center text-[11px] font-bold text-indigo-700 bg-indigo-50 border border-indigo-200 py-1.5 px-3 rounded-xl mb-3 flex items-center justify-center gap-1.5 animate-pulse">
            <span>👉 ဇယားကို ဘယ်/ညာ ပွတ်ဆွဲ (Scroll/Swipe) ၍ အချက်အလက်များအား ပြည့်စုံစွာ ဖတ်ရှုနိုင်ပါသည်</span>
          </div>
        )}

        {/* Paper Container Wrapper with scaling/scrolling control */}
        <div
          className="w-full flex justify-center"
          style={
            viewZoomMode === 'fit' && currentScale < 1
              ? {
                  height: `${paperHeight * currentScale + 16}px`,
                  overflow: 'hidden',
                  width: '100%',
                }
              : {
                  width: '100%',
                  overflowX: 'auto',
                  WebkitOverflowScrolling: 'touch',
                }
          }
        >
          <div
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
            className="transition-transform duration-150 origin-top"
          >
            {/* Paper Container matching exact A4 Landscape */}
            <div
              ref={printRef}
              className="bg-white text-black p-[0.5in] shadow-2xl rounded-sm w-[297mm] min-w-[297mm] border border-slate-300 min-h-[210mm] relative font-pyidaungsu text-[12px] select-text box-border"
            >
          {/* Customizable Title Header */}
          <div className="text-center font-bold text-[14px] text-black font-pyidaungsu mb-1">
            {tableTitle}
          </div>
          <div className="text-right font-bold text-[12px] text-black font-pyidaungsu mb-3 pr-1">
            ရက်စွဲ၊ {useBurmeseDigits ? toBurmeseSlashDate(selectedDate) : formatToDDMMYYYY(selectedDate)}
          </div>

          {/* Data Table matching OTHER ADDRESS RESIDENTS Columns */}
          <table className="w-full text-center text-[11px] border-collapse border border-black font-pyidaungsu table-fixed">
            <colgroup>
              {CRT_COLUMNS_DEFINITION.filter(col => visibleColumns.includes(col.id)).map(col => (
                <col key={col.id} className={col.defaultColWidth} />
              ))}
            </colgroup>
            <thead>
              <tr className="border-b border-black align-middle text-center bg-slate-50">
                {visibleColumns.includes('sr') && (
                  <th 
                    onClick={() => handleHeaderSort('custom')} 
                    className="border border-black p-1.5 font-bold text-[11px] cursor-pointer hover:bg-slate-200 transition-colors"
                  >
                    စဉ်
                  </th>
                )}
                {visibleColumns.includes('passport') && (
                  <th 
                    onClick={() => handleHeaderSort('passport')} 
                    className="border border-black p-1.5 font-bold text-[11px] cursor-pointer hover:bg-slate-200 transition-colors"
                  >
                    နိုင်ငံကူးလက်မှတ်အမှတ်
                  </th>
                )}
                {visibleColumns.includes('nationality') && (
                  <th 
                    onClick={() => handleHeaderSort('nationality')} 
                    className="border border-black p-1.5 font-bold text-[11px] cursor-pointer hover:bg-slate-200 transition-colors"
                  >
                    နိုင်ငံအမည်
                  </th>
                )}
                {visibleColumns.includes('fullname') && (
                  <th 
                    onClick={() => handleHeaderSort('name')} 
                    className="border border-black p-1.5 font-bold text-[11px] cursor-pointer hover:bg-slate-200 transition-colors"
                  >
                    အမည်
                  </th>
                )}
                {visibleColumns.includes('gender') && (
                  <th 
                    className="border border-black p-1.5 font-bold text-[11px]"
                  >
                    ကျား/မ
                  </th>
                )}
                {visibleColumns.includes('visaType') && (
                  <th 
                    onClick={() => handleHeaderSort('visa')} 
                    className="border border-black p-1.5 font-bold text-[11px] cursor-pointer hover:bg-slate-200 transition-colors"
                  >
                    ဗီဇာအမျိုးအစား
                  </th>
                )}
                {visibleColumns.includes('stayPeriod') && (
                  <th 
                    onClick={() => handleHeaderSort('stayTo')} 
                    className="border border-black p-1.5 font-bold text-[11px] cursor-pointer hover:bg-slate-200 transition-colors"
                  >
                    ဗီဇာသက်တမ်း (မှ/ထိ)
                  </th>
                )}
                {visibleColumns.includes('arrivalDate') && (
                  <th 
                    onClick={() => handleHeaderSort('arrival')} 
                    className="border border-black p-1.5 font-bold text-[11px] cursor-pointer hover:bg-slate-200 transition-colors"
                  >
                    နောက်ဆုံးရောက်ရှိရက်စွဲ
                  </th>
                )}
                {visibleColumns.includes('elapsedDays') && (
                  <th 
                    onClick={() => handleHeaderSort('elapsed')} 
                    className="border border-black p-1.5 font-bold text-[11px] cursor-pointer hover:bg-slate-200 transition-colors"
                  >
                    နောက်ဆုံးရောက်ရှိသည့်နေ့မှ ရက်ပေါင်း
                  </th>
                )}
                {visibleColumns.includes('address') && (
                  <th 
                    onClick={() => handleHeaderSort('address')} 
                    className="border border-black p-1.5 font-bold text-[11px] cursor-pointer hover:bg-slate-200 transition-colors"
                  >
                    တည်းခိုလိပ်စာ
                  </th>
                )}
                {visibleColumns.includes('remarks') && (
                  <th 
                    className="border border-black p-1.5 font-bold text-[11px]"
                  >
                    မှတ်ချက်
                  </th>
                )}
              </tr>
            </thead>
            <tbody>
              {displayRows.length === 0 ? (
                <tr>
                  <td colSpan={visibleColumns.length || 1} className="p-6 text-center text-slate-400 font-medium italic border border-black">
                    အချက်အလက် စာရင်းမရှိပါ။
                  </td>
                </tr>
              ) : (
                displayRows.map((m, idx) => {
                  return (
                    <tr key={m.passport} className="border-b border-black">
                      {visibleColumns.includes('sr') && (
                        <td className="p-1.5 border border-black text-center font-bold">
                          {useBurmeseDigits ? toBurmeseDigits(idx + 1) : idx + 1}
                        </td>
                      )}
                      {visibleColumns.includes('passport') && (
                        <td className="p-1.5 border border-black text-center font-bold font-mono break-words">
                          {m.passport}
                        </td>
                      )}
                      {visibleColumns.includes('nationality') && (
                        <td className="p-1.5 border border-black text-center font-bold break-words">
                          {m.nationality}
                        </td>
                      )}
                      {visibleColumns.includes('fullname') && (
                        <td className="p-1.5 border border-black text-left font-bold break-words">
                          {m.fullname}
                        </td>
                      )}
                      {visibleColumns.includes('gender') && (
                        <td className="p-1.5 border border-black text-center">
                          {m.gender === 'M' ? 'ကျား' : m.gender === 'F' ? 'မ' : '-'}
                        </td>
                      )}
                      {visibleColumns.includes('visaType') && (
                        <td className="p-1.5 border border-black text-center break-words">
                          {m.visaType}
                        </td>
                      )}
                      {visibleColumns.includes('stayPeriod') && (
                        <td className="p-1.5 border border-black text-center break-words">
                          {(m.stayFrom || m.stayTo) ? `${m.stayFrom || '-'} မှ ${m.stayTo || '-'}` : '-'}
                        </td>
                      )}
                      {visibleColumns.includes('arrivalDate') && (
                        <td className="p-1.5 border border-black text-center">
                          {m.lastArrivalDate}
                        </td>
                      )}
                      {visibleColumns.includes('elapsedDays') && (
                        <td className="p-1.5 border border-black text-center font-bold">
                          {useBurmeseDigits ? toBurmeseDigits(m.elapsedDays) : m.elapsedDays}
                        </td>
                      )}
                      {visibleColumns.includes('address') && (
                        <td className="p-1.5 border border-black text-left font-bold break-words">
                          {m.address}
                        </td>
                      )}
                      {visibleColumns.includes('remarks') && (
                        <td className="p-1.5 border border-black text-left break-words text-[10px]">
                          {m.remarks || ''}
                        </td>
                      )}
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>

          {/* Officer Signature Footer */}
          {showSigner && (
            <div className="flex justify-end mt-12 pr-6">
              <div className="text-center font-pyidaungsu flex flex-col items-center min-w-[280px] p-3 rounded-2xl border border-transparent hover:border-dashed hover:border-indigo-300 transition-all bg-transparent hover:bg-indigo-50/20 group">
                <div className="h-10" />
                <div className="w-full">
                  <input
                    type="text"
                    value={officerTitle}
                    onChange={(e) => handleSaveOfficerTitle(e.target.value)}
                    placeholder="ရာထူး ရိုက်ထည့်ပါ..."
                    list="crtOfficialTitlesList"
                    className="w-full text-center text-[13px] font-extrabold text-black bg-transparent border-b border-transparent group-hover:border-slate-300 focus:border-indigo-500 focus:bg-white focus:shadow-xs outline-none py-0.5 transition-all font-pyidaungsu"
                    title="လက်မှတ်ထိုးသူ ရာထူး (Click to edit)"
                  />
                </div>
                <div className="w-full mt-1">
                  <input
                    type="text"
                    value={officerName}
                    onChange={(e) => handleSaveOfficerName(e.target.value)}
                    placeholder="အမည် ရိုက်ထည့်ပါ..."
                    list="crtOfficialNamesList"
                    className="w-full text-center text-[13px] font-bold text-black bg-transparent border-b border-transparent group-hover:border-slate-300 focus:border-indigo-500 focus:bg-white focus:shadow-xs outline-none py-0.5 transition-all font-pyidaungsu"
                    title="လက်မှတ်ထိုးသူ အမည် (Click to edit)"
                  />
                </div>
                <div className="print:hidden text-[9px] text-slate-400 font-sans mt-1 opacity-0 group-hover:opacity-100 transition-opacity">
                  ✏️ လက်မှတ် ရာထူး/အမည် စိတ်ကြိုက်ပြင်ရန် နှိပ်ပါ
                </div>
              </div>
            </div>
          )}
        </div>
          </div>
        </div>
      </div>

      {/* Column Visibility Settings Modal */}
      <CustomReportColumnModal
        isOpen={isColumnModalOpen}
        onClose={() => setIsColumnModalOpen(false)}
        visibleColumns={visibleColumns}
        onToggleColumn={handleToggleColumn}
        onSelectAll={handleSelectAllColumns}
        onResetDefault={handleResetDefaultColumns}
      />

      {/* Report Template Presets Modal */}
      <CustomReportPresetsModal
        isOpen={isPresetModalOpen}
        onClose={() => setIsPresetModalOpen(false)}
        currentSettings={{
          tableTitle,
          showSigner,
          officerTitle,
          officerName,
          useBurmeseDigits,
          dataPool,
          visibleColumns,
          filterNationalities,
          filterVisaTypes,
          filterAddresses,
          filterAddressType,
          filterGender,
          minElapsedDays
        }}
        onApplyPreset={handleApplyPreset}
        showToast={showToast}
      />
    </div>
  );
};
