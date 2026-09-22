import { MasterItem } from '../types';

/**
 * Standardized helper to determine whether an address represents a Company (Co., Ltd)
 * vs Other (Hotel, Guesthouse, Residence, General).
 * Handles English variations, Myanmar text, and linked Master DB categories.
 */
export const isCompanyAddress = (addressStr?: string | null, masterData?: MasterItem[]): boolean => {
  if (!addressStr) return false;
  const clean = String(addressStr).trim();
  if (!clean) return false;
  const upper = clean.toUpperCase();

  // 1. Check masterData if linked
  if (masterData && masterData.length > 0) {
    const matchedMaster = masterData.find(m => 
      m && m.type === 'Stay' && m.name && (
        m.name.trim().toLowerCase() === clean.toLowerCase() ||
        clean.toLowerCase().includes(m.name.trim().toLowerCase())
      )
    );
    if (matchedMaster) {
      const masterName = (matchedMaster.name || '').toUpperCase();
      const masterLinked = (matchedMaster.linkedValue || '').toUpperCase();
      if (
        (masterName.includes('CO') && (masterName.includes('LTD') || masterName.includes('L.T.D') || masterName.includes('LIMITED'))) ||
        masterName.includes('COMPANY') ||
        masterName.includes('CORP') ||
        masterName.includes('ENTERPRISE') ||
        masterName.includes('ကုမ္ပဏီ') ||
        masterName.includes('လီမိတက်') ||
        (masterLinked.includes('CO') && (masterLinked.includes('LTD') || masterLinked.includes('L.T.D') || masterLinked.includes('LIMITED'))) ||
        masterLinked.includes('COMPANY') ||
        masterLinked.includes('ကုမ္ပဏီ')
      ) {
        return true;
      }
    }
  }

  // 2. Check string keywords (English and Myanmar)
  if (
    (upper.includes('CO') && (upper.includes('LTD') || upper.includes('L.T.D') || upper.includes('LIMITED'))) ||
    upper.includes('CO.,') ||
    upper.includes('CO,') ||
    upper.includes('CO. ') ||
    upper.includes('CO ') ||
    upper.includes('COMPANY') ||
    upper.includes('CORP') ||
    upper.includes('ENTERPRISE') ||
    upper.includes('LTD') ||
    upper.includes('L.T.D') ||
    upper.includes('LIMITED') ||
    clean.includes('ကုမ္ပဏီ') ||
    clean.includes('ကုမ္မဏီ') ||
    clean.includes('လီမိတက်')
  ) {
    return true;
  }

  return false;
};
