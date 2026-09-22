import { CRTColumnId, CRTPreset } from '../types';

export interface CRTColumnDef {
  id: CRTColumnId;
  label: string;
  shortLabel: string;
  defaultColWidth: string; // e.g. "w-[30px]"
  excelWidth: number;
}

export const CRT_COLUMNS_DEFINITION: CRTColumnDef[] = [
  { id: 'sr', label: 'စဉ်', shortLabel: 'စဉ်', defaultColWidth: 'w-[30px]', excelWidth: 6 },
  { id: 'passport', label: 'နိုင်ငံကူးလက်မှတ်အမှတ်', shortLabel: 'ပတ်စပို့', defaultColWidth: 'w-[85px]', excelWidth: 18 },
  { id: 'nationality', label: 'နိုင်ငံအမည်', shortLabel: 'နိုင်ငံ', defaultColWidth: 'w-[50px]', excelWidth: 15 },
  { id: 'fullname', label: 'အမည်', shortLabel: 'အမည်', defaultColWidth: 'w-[115px]', excelWidth: 25 },
  { id: 'gender', label: 'ကျား/မ', shortLabel: 'ကျား/မ', defaultColWidth: 'w-[40px]', excelWidth: 8 },
  { id: 'visaType', label: 'ဗီဇာအမျိုးအစား', shortLabel: 'ဗီဇာ', defaultColWidth: 'w-[80px]', excelWidth: 18 },
  { id: 'stayPeriod', label: 'ဗီဇာသက်တမ်း (မှ/ထိ)', shortLabel: 'သက်တမ်း', defaultColWidth: 'w-[105px]', excelWidth: 25 },
  { id: 'arrivalDate', label: 'နောက်ဆုံးရောက်ရှိရက်စွဲ', shortLabel: 'ရောက်ရှိရက်', defaultColWidth: 'w-[80px]', excelWidth: 18 },
  { id: 'elapsedDays', label: 'နောက်ဆုံးရောက်ရှိသည့်နေ့မှ ရက်ပေါင်း', shortLabel: 'ရက်ပေါင်း', defaultColWidth: 'w-[65px]', excelWidth: 16 },
  { id: 'address', label: 'တည်းခိုလိပ်စာ', shortLabel: 'လိပ်စာ', defaultColWidth: 'w-[140px]', excelWidth: 30 },
  { id: 'remarks', label: 'မှတ်ချက်', shortLabel: 'မှတ်ချက်', defaultColWidth: 'w-[125px]', excelWidth: 20 },
];

export const ALL_COLUMN_IDS: CRTColumnId[] = CRT_COLUMNS_DEFINITION.map(c => c.id);

export const DEFAULT_BUILT_IN_PRESETS: CRTPreset[] = [
  {
    id: 'preset_other_address',
    name: 'အခြားလိပ်စာနေထိုင်သူများစာရင်း',
    isBuiltIn: true,
    tableTitle: 'လက်ရှိနေထိုင်သူများမှ အခြားလိပ်စာဖြင့်နေထိုင်သူများ',
    officerTitle: 'လဝကမှူး',
    officerName: 'ဒုတိယလဝကမှူး',
    useBurmeseDigits: true,
    dataPool: 'STILL_IN',
    visibleColumns: [...ALL_COLUMN_IDS],
    filterAddressType: 'OTHER',
    description: 'အခြားလိပ်စာဖြင့် တည်းခိုနေထိုင်သူများအားလုံးအတွက် စံပြ အစီရင်ခံစာ ပုံစံ'
  },
  {
    id: 'preset_hotel_stay',
    name: 'ဟိုတယ်တည်းခိုသူများစာရင်း',
    isBuiltIn: true,
    tableTitle: 'လက်ရှိနေထိုင်သူများမှ ဟိုတယ်/တည်းခိုခန်းများတွင် တည်းခိုသူများ',
    officerTitle: 'လဝကမှူး',
    officerName: 'ဒုတိယလဝကမှူး',
    useBurmeseDigits: true,
    dataPool: 'STILL_IN',
    visibleColumns: [...ALL_COLUMN_IDS],
    filterAddress: 'ALL',
    description: 'ဟိုတယ်နှင့် တည်းခိုခန်းများတွင် တည်းခိုနေထိုင်သူများ အစီရင်ခံစာ'
  },
  {
    id: 'preset_long_stay_overstay',
    name: 'ရက်ပေါင်း ၃၀ ကျော် နေထိုင်သူများ',
    isBuiltIn: true,
    tableTitle: 'မြိတ်မြို့နယ်အတွင်း ရက်ပေါင်း (၃၀) အထက် ကာလရှည်နေထိုင်သူများစာရင်း',
    officerTitle: 'လဝကမှူး',
    officerName: 'ဒုတိယလဝကမှူး',
    useBurmeseDigits: true,
    dataPool: 'STILL_IN',
    visibleColumns: ['sr', 'passport', 'nationality', 'fullname', 'gender', 'visaType', 'stayPeriod', 'arrivalDate', 'elapsedDays', 'address', 'remarks'],
    minElapsedDays: '30',
    description: 'ရက်ပေါင်း ၃၀ အထက် ကာလရှည်နေထိုင်သူများ စာရင်း'
  },
  {
    id: 'preset_inbound_daily',
    name: 'ဝင်ရောက်လာသူများ အစီရင်ခံစာ (Inbound)',
    isBuiltIn: true,
    tableTitle: 'လေကြောင်းလိုင်းဖြင့် ဝင်ရောက်လာသော နိုင်ငံခြားသားများစာရင်း',
    officerTitle: 'လဝကမှူး',
    officerName: 'ဒုတိယလဝကမှူး',
    useBurmeseDigits: true,
    dataPool: 'INBOUND',
    visibleColumns: ['sr', 'passport', 'nationality', 'fullname', 'gender', 'visaType', 'stayPeriod', 'arrivalDate', 'address', 'remarks'],
    description: 'ဝင်ရောက်လာသော နိုင်ငံခြားသားများ စာရင်း'
  },
  {
    id: 'preset_compact_summary',
    name: 'အကျဉ်းချုပ် ဇယား (Compact View)',
    isBuiltIn: true,
    tableTitle: 'လက်ရှိနေထိုင်သူများ စာရင်း အကျဉ်းချုပ်',
    officerTitle: 'လဝကမှူး',
    officerName: 'ဒုတိယလဝကမှူး',
    useBurmeseDigits: true,
    dataPool: 'STILL_IN',
    visibleColumns: ['sr', 'passport', 'nationality', 'fullname', 'visaType', 'elapsedDays', 'address', 'remarks'],
    description: 'အရေးကြီး ကော်လံများသာ ပါဝင်သော အကျဉ်းချုပ် ပုံစံ'
  }
];

const CUSTOM_PRESETS_KEY = 'crt_custom_presets_v2';

export function getStoredCustomPresets(): CRTPreset[] {
  try {
    const raw = localStorage.getItem(CUSTOM_PRESETS_KEY);
    if (raw) {
      return JSON.parse(raw);
    }
  } catch (e) {
    console.error("Failed to load custom presets:", e);
  }
  return [];
}

export function saveCustomPreset(preset: Omit<CRTPreset, 'id' | 'createdAt'>): CRTPreset {
  const customPresets = getStoredCustomPresets();
  const newPreset: CRTPreset = {
    ...preset,
    id: `PRESET-${Date.now()}`,
    isBuiltIn: false,
    createdAt: new Date().toISOString()
  };
  const updated = [newPreset, ...customPresets];
  try {
    localStorage.setItem(CUSTOM_PRESETS_KEY, JSON.stringify(updated));
  } catch (e) {
    console.error("Failed to save custom preset:", e);
  }
  return newPreset;
}

export function deleteCustomPreset(id: string): void {
  const customPresets = getStoredCustomPresets();
  const updated = customPresets.filter(p => p.id !== id);
  try {
    localStorage.setItem(CUSTOM_PRESETS_KEY, JSON.stringify(updated));
  } catch (e) {
    console.error("Failed to delete custom preset:", e);
  }
}
