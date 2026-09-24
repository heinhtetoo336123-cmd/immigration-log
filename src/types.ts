export type Mode = 'IN' | 'OUT';

export type PermitStatus = 'နေထိုင်ခွင့်ကျထားသောသူ' | 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ';

export const normalizePermitStatus = (status?: string | null): PermitStatus | '' => {
  if (!status) return '';
  const s = String(status).trim();
  if (
    s === 'နေထိုင်ခွင့်ကျထားသောသူ' ||
    s === 'STAY PERMITTED' ||
    s === 'STILL PERMITTED' ||
    s === 'STILL CONFIRMED' ||
    s === 'CONFIRMED' ||
    s.toLowerCase().includes('ကျထား')
  ) {
    return 'နေထိုင်ခွင့်ကျထားသောသူ';
  }
  if (
    s === 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ' ||
    s === 'STAY NOT PERMITTED' ||
    s === 'STILL NOT PERMITTED YET' ||
    s === 'NOT PERMITTED' ||
    s.toLowerCase().includes('မလျှောက်') ||
    s.toLowerCase().includes('မလျောက်')
  ) {
    return 'နေထိုင်ခွင့် မလျှောက်ထားသေးသူ';
  }
  return s as PermitStatus;
};

export interface ImmRecord {
  id: number;
  timestamp: string;
  mode: Mode;
  logType?: 'FFE' | 'FCR';
  passport: string;
  fullname: string;
  gender: 'M' | 'F' | '';
  dob?: string;
  nationality: string;
  address: string;
  stayDescription?: string;
  visaType: string;
  visaNumber?: string;
  vehicleInfo: string;
  stayFrom: string;
  stayTo: string;
  totalDays: string;
  remainingDays?: string;
  arrivedFrom: string;
  departedTo: string;
  broughtBy: string;
  contactDetails: string;
  officialName?: string;
  officialTitle?: string;
  remarks?: string;
  stillPermittedStatus?: PermitStatus | string;
  permittedBy?: string;
  previousPassport?: string;
  dualPassportRemarks?: string;
  linkedPassports?: string[];
  formC?: any;
  syncStatus?: 'pending_sync' | 'upload_failed' | 'synced';
  updatedAt?: string;
  serverSyncedAt?: string;
  createdDevice?: string;
}

export interface MasterItem {
  id: number;
  name: string;
  type: 'Nationality' | 'Visa' | 'Stay' | 'Vehicle' | 'Agent' | 'Contact' | 'Official' | 'Title' | 'Reporter' | 'Phone' | 'PermitDescription';
  linkedValue?: string;
  updatedAt?: string;
  syncStatus?: 'pending_sync' | 'upload_failed' | 'synced';
  serverSyncedAt?: string;
}

export type Tab = 'entry' | 'data' | 'movement' | 'stillIn' | 'daily' | 'master' | 'tableOutput' | 'customReport' | 'checking' | 'individualSearch' | 'alarm' | 'telegraph' | 'watchList' | 'info' | 'auditLog';

export interface WatchListPerson {
  id: string;
  name: string;
  passport: string;
  nationality: string;
  fatherName?: string;
  motherName?: string;
  dob?: string;
  birthPlace?: string;
  address?: string;
  idNumber?: string;
  idType?: string;
  additionalInfo?: string;
}

export interface OfficialConfirmation {
  status: 'CONFIRMED_MATCH' | 'CLEARED_MISMATCH' | 'PENDING';
  confirmedBy?: string;
  confirmedTitle?: string;
  confirmedAt?: string;
  remarks?: string;
  personDecisions?: Record<string, {
    decision: 'CONFIRMED_MATCH' | 'CLEARED_MISMATCH' | 'PENDING';
    remarks?: string;
    confirmedBy?: string;
    confirmedAt?: string;
  }>;
}

export interface WatchListRecord {
  id: string;
  letterDate: string;
  letterNo: string;
  requestDepartment: string;
  reasonDescription: string;
  persons: WatchListPerson[];
  totalPersons?: number;
  officerName?: string;
  officerTitle?: string;
  createdAt: string;
  status?: 'PENDING' | 'CHECKED' | 'FLAGGED' | 'CONFIRMED_MATCH' | 'CLEARED_MISMATCH';
  lastCheckedAt?: string;
  findingsSummary?: string;
  matchResults?: any;
  officialConfirmation?: OfficialConfirmation;
}


export interface DeviceSession {
  deviceId: string;
  deviceName: string;
  accountRole: 'Superadmin' | 'Editor';
  username: string;
  lastActive: string;
  kicked?: boolean;
  readCount?: number;
  writeCount?: number;
}

export interface InvNote {
  id: string;
  passport?: string;
  authorName: string;
  authorTitle: string;
  timestamp: string;
  content: string;
}

export interface VehicleSummary {
  id: number;
  date: string;
  vehicleNo: string;
  time: string;
  totalInM: number;
  totalInT: number;
  totalOutM: number;
  totalOutT: number;
  remark: string;
  updatedAt?: string;
  syncStatus?: 'pending_sync' | 'upload_failed' | 'synced';
}

export interface MovementData {
  p: string;
  n: string;
  nat: string;
  loc: string;
  visa: string;
  visaNumber?: string;
  dob?: string;
  gender: 'M' | 'F' | '';
  start: string;
  end: string;
  allowed: string;
  vInfo: string;
  agent: string;
  contact: string;
  offName: string;
  offTitle: string;
  repName: string;
  repPhone: string;
  in: string;
  out: string;
  inTime: number;
  outTime: number;
  latestTime: number;
  lastId: number;
  previousPassport?: string;
  dualPassportRemarks?: string;
  linkedPassports?: string[];
  stayDescription?: string;
  stillPermittedStatus?: string;
  permittedBy?: string;
  remarks?: string;
  isStillIn?: boolean;
}

export interface CheckingHistoryEntry {
  id: string;
  passport: string;
  fullname: string;
  nationality: string;
  type: 'CO_LTD' | 'OTHERS';
  checkDate: string; // YYYY-MM-DD
  originalAddress: string;
  confirmedAddress: string;
  confirmedStayDescription?: string;
  status: PermitStatus | string;
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

export interface DossierRecord {
  id: string; // e.g. IMD-YYYYMMDD-XXXX
  generatedAt: string;
  passports: string[];
  fileType: string;
  dossierTitle: string;
  remarksMap?: Record<string, string>;
}

export type CRTColumnId = 
  | 'sr'
  | 'passport'
  | 'nationality'
  | 'fullname'
  | 'gender'
  | 'visaType'
  | 'stayPeriod'
  | 'arrivalDate'
  | 'elapsedDays'
  | 'address'
  | 'remarks';

export interface CRTPreset {
  id: string;
  name: string;
  isBuiltIn?: boolean;
  tableTitle: string;
  officerTitle: string;
  officerName: string;
  showSigner?: boolean;
  useBurmeseDigits: boolean;
  dataPool: 'STILL_IN' | 'ALL_MOVEMENTS' | 'INBOUND' | 'OUTBOUND';
  visibleColumns: CRTColumnId[];
  filterNationality?: string;
  filterNationalities?: string[];
  filterVisaType?: string;
  filterVisaTypes?: string[];
  filterAddress?: string;
  filterAddresses?: string[];
  filterAddressType?: 'ALL' | 'COMPANY' | 'OTHER';
  filterGender?: 'ALL' | 'M' | 'F';
  minElapsedDays?: string;
  description?: string;
  createdAt?: string;
}

export type ActivityActionType = 
  | 'CREATE' 
  | 'UPDATE' 
  | 'DELETE' 
  | 'BACKUP' 
  | 'RESTORE' 
  | 'SYNC' 
  | 'WATCHLIST_MATCH' 
  | 'SESSION_KICK' 
  | 'RESET_QUOTA'
  | 'SETTINGS_CHANGE';

export type ActivityModule = 
  | 'FFE' 
  | 'FCR' 
  | 'CHECKING' 
  | 'MASTER' 
  | 'WATCHLIST' 
  | 'SYSTEM' 
  | 'CRT' 
  | 'TELEG'
  | 'CLOUD_AUTH';

export interface ActivityLogEntry {
  id: string;
  timestamp: string; // ISO String
  readableTime: string; // e.g. DD/MM/YYYY, HH:mm:ss
  action: ActivityActionType;
  module: ActivityModule;
  officerName: string;
  officerRole: string;
  deviceId?: string;
  targetId?: string; // e.g. Passport number or ID
  details: string;
}

