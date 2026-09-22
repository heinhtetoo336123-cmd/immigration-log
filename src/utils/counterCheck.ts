import { ImmRecord } from '../types';

export interface SuspectGroup {
  id: string;
  type: 
    | 'SAME_NAME_DOB_DIFF_PASSPORT'  // Exact match on Name & DOB, different Passports (Strong Dual/Renewed Passport or Typo)
    | 'SAME_PASSPORT_DIFF_DOB'       // Exact Passport, but conflicting DOBs (Critical Entry Error / Mixed Profile)
    | 'SAME_NAME_DIFF_PASSPORT'      // Name match with no/varying DOB
    | 'SAME_PASSPORT_DIFF_NAME'      // Same Passport, different Names
    | 'SIMILAR_PASSPORT_TYPO'        // 1-2 character typo with similar/same person
    | 'DUAL_PASSPORT_LINKED';        // Confirmed Dual / Renewed Passport chain
  title: string;
  description: string;
  badgeColor: string;
  records: ImmRecord[];
  suggestedPassport: string;
  suggestedName: string;
  suggestedDob?: string;
  distinctPassports: string[];
  distinctNames: string[];
  distinctDobs: string[];
  distinctNationalities?: string[];
  isAlreadyLinked?: boolean;
  linkedPassports?: string[];
  dualPassportRemarks?: string;
  confidenceScore?: number; // 0 - 100%
  reasonNote?: string;
}

// 1. Name Normalization
export function normalizeName(name: string): string {
  if (!name) return '';
  let clean = name.toUpperCase().trim();
  // Remove common prefixes/titles
  clean = clean.replace(/\b(MR|MRS|MS|MISS|DR|DAW|U|MH|MD|MASTER)\b/gi, ' ');
  // Replace punctuation with space
  clean = clean.replace(/[^A-Z0-9\s]/gi, ' ');
  // Collapse multiple spaces
  clean = clean.replace(/\s+/g, ' ').trim();
  return clean;
}

// 2. Tokenized Sorted Name for word order independence (e.g. "JOHN DOE" vs "DOE JOHN")
export function tokenSortedName(name: string): string {
  const norm = normalizeName(name);
  if (!norm) return '';
  return norm.split(' ').filter(Boolean).sort().join(' ');
}

// 3. Passport Normalization
export function normalizePassport(passport: string): string {
  if (!passport) return '';
  return passport.toUpperCase().replace(/[^A-Z0-9]/gi, '').trim();
}

// 4. DOB Normalization (Standardize YYYY-MM-DD, DD/MM/YYYY, DD-MM-YYYY)
export function normalizeDob(dob?: string): string {
  if (!dob) return '';
  const clean = dob.trim();
  if (!clean) return '';

  // Match YYYY-MM-DD or YYYY/MM/DD
  const isoMatch = clean.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (isoMatch) {
    const y = isoMatch[1];
    const m = isoMatch[2].padStart(2, '0');
    const d = isoMatch[3].padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  // Match DD-MM-YYYY or DD/MM/YYYY
  const dmyMatch = clean.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (dmyMatch) {
    const d = dmyMatch[1].padStart(2, '0');
    const m = dmyMatch[2].padStart(2, '0');
    const y = dmyMatch[3];
    return `${y}-${m}-${d}`;
  }

  return clean;
}

// 5. Levenshtein Distance
export function levenshteinDistance(a: string, b: string): number {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;

  const matrix: number[][] = [];
  for (let i = 0; i <= b.length; i++) {
    matrix[i] = [i];
  }
  for (let j = 0; j <= a.length; j++) {
    matrix[0][j] = j;
  }

  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // substitution
          matrix[i][j - 1] + 1,     // insertion
          matrix[i - 1][j] + 1      // deletion
        );
      }
    }
  }
  return matrix[b.length][a.length];
}

// Check OCR / Key confusion similarity
export function isSimilarPassport(p1: string, p2: string): boolean {
  if (p1 === p2) return false;
  const lenDiff = Math.abs(p1.length - p2.length);
  if (lenDiff > 2) return false;

  const dist = levenshteinDistance(p1, p2);
  if (dist <= 1) return true;
  if (dist === 2 && Math.min(p1.length, p2.length) >= 7) return true;

  // OCR character replacements check (0<->O, 1<->I, 8<->B, 5<->S)
  const p1Sub = p1.replace(/0/g, 'O').replace(/1/g, 'I').replace(/8/g, 'B').replace(/5/g, 'S');
  const p2Sub = p2.replace(/0/g, 'O').replace(/1/g, 'I').replace(/8/g, 'B').replace(/5/g, 'S');
  if (p1Sub === p2Sub) return true;

  return false;
}

/**
 * Builds a Unified Passport Chain Map for all records.
 * If Passport A is linked to Passport B, both map to [A, B].
 */
export function buildPersonPassportChainMap(records: ImmRecord[]): Map<string, string[]> {
  const chainMap = new Map<string, Set<string>>();

  const addEdge = (p1: string, p2: string) => {
    const np1 = normalizePassport(p1);
    const np2 = normalizePassport(p2);
    if (!np1 || !np2 || np1 === np2) return;

    if (!chainMap.has(np1)) chainMap.set(np1, new Set([np1]));
    if (!chainMap.has(np2)) chainMap.set(np2, new Set([np2]));

    const set1 = chainMap.get(np1)!;
    const set2 = chainMap.get(np2)!;
    const union = new Set([...set1, ...set2]);

    union.forEach(p => {
      chainMap.set(p, union);
    });
  };

  (records || []).forEach(r => {
    const curP = normalizePassport(r.passport);
    if (!curP) return;

    if (!chainMap.has(curP)) chainMap.set(curP, new Set([curP]));

    if (r.previousPassport) {
      addEdge(curP, r.previousPassport);
    }
    if (Array.isArray(r.linkedPassports)) {
      r.linkedPassports.forEach(lp => addEdge(curP, lp));
    }
  });

  const result = new Map<string, string[]>();
  chainMap.forEach((set, key) => {
    result.set(key, Array.from(set));
  });

  return result;
}

/**
 * Returns all passports linked to the given passport in the foreigner profile chain.
 */
export function getLinkedPassportsForPassport(records: ImmRecord[], passport: string): string[] {
  const np = normalizePassport(passport);
  if (!np) return [];
  const chainMap = buildPersonPassportChainMap(records);
  const found = chainMap.get(np);
  if (found && found.length > 0) return found;
  return [np];
}

/**
 * Enhanced Counter Check Analyzer incorporating complete FFE fields including DOB:
 * 1. Category 0: Confirmed Dual / Renewed Passports
 * 2. Category 1: High-Confidence Dual Passport / Identity (SAME NAME + SAME DOB, DIFFERENT PASSPORT)
 * 3. Category 2: Data Entry Conflict (SAME PASSPORT, CONFLICTING DOBs)
 * 4. Category 3: Same Name, Different Passports (DOB missing or varying)
 * 5. Category 4: Same Passport, Different Names
 * 6. Category 5: Similar Passport Typo (Levenshtein match with person correlation)
 */
export function analyzeRecordsForErrors(records: ImmRecord[]): SuspectGroup[] {
  const suspectGroups: SuspectGroup[] = [];
  const processedGroupKeys = new Set<string>();

  // Filter valid records
  const validRecords = (records || []).filter(r => r.passport && r.passport.trim() && r.fullname && r.fullname.trim());

  // --- CATEGORY 0: EXPLICITLY LINKED DUAL / RENEWED PASSPORTS ---
  const linkedChainMap = buildPersonPassportChainMap(records);
  const processedChains = new Set<string>();

  linkedChainMap.forEach((chain) => {
    if (chain.length > 1) {
      const sortedChainKey = [...chain].sort().join('_');
      if (!processedChains.has(sortedChainKey)) {
        processedChains.add(sortedChainKey);
        
        const chainRecords = validRecords.filter(r => chain.includes(normalizePassport(r.passport)));
        if (chainRecords.length > 0) {
          const distinctPList = Array.from(new Set(chainRecords.map(r => r.passport.toUpperCase().trim())));
          const distinctNList = Array.from(new Set(chainRecords.map(r => r.fullname.trim())));
          const distinctDList = Array.from(new Set(chainRecords.map(r => r.dob ? normalizeDob(r.dob) : '').filter(Boolean)));
          const distinctNatList = Array.from(new Set(chainRecords.map(r => r.nationality || '').filter(Boolean)));

          const primaryP = distinctPList[0] || chain[0];
          const primaryN = distinctNList[0] || '';
          const primaryDob = distinctDList[0] || '';
          const existingRemark = chainRecords.find(r => r.dualPassportRemarks)?.dualPassportRemarks || '';

          const groupKey = `DUAL_CHAIN_${sortedChainKey}`;
          processedGroupKeys.add(groupKey);
          suspectGroups.push({
            id: groupKey,
            type: 'DUAL_PASSPORT_LINKED',
            title: `🔗 အတည်ပြုပြီး Dual / Renewed Passports (${distinctPList.join(' ⇄ ')})`,
            description: `ပုဂ္ဂိုလ် (${primaryN}${primaryDob ? ` / မွေးသက္ကရာဇ်: ${primaryDob}` : ''}) သည် Passport (${distinctPList.join(', ')}) နံပါတ်များ ချိတ်ဆက်ထားသော အတည်ပြုပြီး Foreigner Profile ဖြစ်ပါသည်။`,
            badgeColor: 'bg-emerald-600 text-white',
            records: chainRecords,
            suggestedPassport: primaryP,
            suggestedName: primaryN,
            suggestedDob: primaryDob,
            distinctPassports: distinctPList,
            distinctNames: distinctNList,
            distinctDobs: distinctDList,
            distinctNationalities: distinctNatList,
            isAlreadyLinked: true,
            linkedPassports: distinctPList,
            dualPassportRemarks: existingRemark,
            confidenceScore: 100,
            reasonNote: 'ကြိုတင်ချိတ်ဆက်ထားပြီးဖြစ်သော Unified Passport Profile'
          });
        }
      }
    }
  });

  // --- CATEGORY 1: SAME NAME & SAME DOB, BUT DIFFERENT PASSPORTS ---
  // High confidence indicator that this is the same individual holding renewed/dual passports or a typo!
  const nameDobToRecordsMap = new Map<string, ImmRecord[]>();
  for (const r of validRecords) {
    const sName = tokenSortedName(r.fullname);
    const nDob = normalizeDob(r.dob);
    if (!sName || sName.length < 3 || !nDob) continue;

    const key = `${sName}__DOB__${nDob}`;
    if (!nameDobToRecordsMap.has(key)) {
      nameDobToRecordsMap.set(key, []);
    }
    nameDobToRecordsMap.get(key)!.push(r);
  }

  nameDobToRecordsMap.forEach((groupRecords, key) => {
    const distinctPList = Array.from(new Set(groupRecords.map(r => r.passport.toUpperCase().trim())));
    if (distinctPList.length > 1) {
      const distinctNList = Array.from(new Set(groupRecords.map(r => r.fullname.trim())));
      const distinctDList = Array.from(new Set(groupRecords.map(r => normalizeDob(r.dob)).filter(Boolean)));
      const distinctNatList = Array.from(new Set(groupRecords.map(r => r.nationality || '').filter(Boolean)));

      // Check if already completely linked
      const isAlreadyChain = distinctPList.every(p => {
        const chain = linkedChainMap.get(normalizePassport(p));
        return chain && distinctPList.every(other => chain.includes(normalizePassport(other)));
      });

      const passportCountMap = new Map<string, number>();
      groupRecords.forEach(r => {
        const np = normalizePassport(r.passport);
        passportCountMap.set(np, (passportCountMap.get(np) || 0) + 1);
      });
      let maxPCount = 0;
      let suggestedP = distinctPList[0];
      passportCountMap.forEach((count, p) => {
        if (count > maxPCount) {
          maxPCount = count;
          const found = groupRecords.find(r => normalizePassport(r.passport) === p);
          if (found) suggestedP = found.passport;
        }
      });

      const groupKey = `SAME_NAME_DOB_${key}_${distinctPList.sort().join('_')}`;
      if (!processedGroupKeys.has(groupKey)) {
        processedGroupKeys.add(groupKey);
        suspectGroups.push({
          id: groupKey,
          type: isAlreadyChain ? 'DUAL_PASSPORT_LINKED' : 'SAME_NAME_DOB_DIFF_PASSPORT',
          title: isAlreadyChain
            ? `🔗 အတည်ပြုပြီး Dual / Renewed Passports (${distinctPList.join(' ⇄ ')})`
            : `⭐ အမည်နှင့် DOB ထပ်တူညီသော်လည်း Passport မတူပါ (${distinctNList[0]})`,
          description: isAlreadyChain
            ? `ပုဂ္ဂိုလ် (${distinctNList[0]}) ၏ Passport များ ချိတ်ဆက်ပြီးဖြစ်ပါသည်။`
            : `အမည် '${distinctNList[0]}' နှင့် မွေးသက္ကရာဇ် (DOB: ${distinctDList[0] || '-'}) အတိအကျ ထပ်တူညီနေသော်လည်း Passport နံပါတ် (${distinctPList.join(', ')}) ကွဲပြားနေပါသည်။ ဤသည်မှာ ၉၉% သေချာသော Passport အသစ်လဲခြင်း (Renewed/Dual Passport) သို့မဟုတ် စာရိုက်မှားယွင်းမှု ဖြစ်ပါသည်။`,
          badgeColor: isAlreadyChain ? 'bg-emerald-600 text-white' : 'bg-purple-600 text-white',
          records: groupRecords,
          suggestedPassport: suggestedP,
          suggestedName: distinctNList[0],
          suggestedDob: distinctDList[0],
          distinctPassports: distinctPList,
          distinctNames: distinctNList,
          distinctDobs: distinctDList,
          distinctNationalities: distinctNatList,
          isAlreadyLinked: isAlreadyChain,
          linkedPassports: distinctPList,
          confidenceScore: 98,
          reasonNote: `Full Name (${distinctNList[0]}) + DOB (${distinctDList[0]}) Match`
        });
      }
    }
  });

  // --- CATEGORY 2: SAME PASSPORT, BUT CONFLICTING DOBS ---
  // If one passport has conflicting DOB values (e.g. 1990-01-01 vs 1995-05-12), that's a high risk entry error!
  const passportToRecordsMap = new Map<string, ImmRecord[]>();
  for (const r of validRecords) {
    const np = normalizePassport(r.passport);
    if (!np || np.length < 4) continue;
    if (!passportToRecordsMap.has(np)) {
      passportToRecordsMap.set(np, []);
    }
    passportToRecordsMap.get(np)!.push(r);
  }

  passportToRecordsMap.forEach((groupRecords, normPassportKey) => {
    const distinctDobs = Array.from(new Set(groupRecords.map(r => normalizeDob(r.dob)).filter(Boolean)));
    if (distinctDobs.length > 1) {
      const distinctPList = Array.from(new Set(groupRecords.map(r => r.passport.toUpperCase().trim())));
      const distinctNList = Array.from(new Set(groupRecords.map(r => r.fullname.trim())));
      const distinctNatList = Array.from(new Set(groupRecords.map(r => r.nationality || '').filter(Boolean)));

      // Find most frequent DOB
      const dobCount = new Map<string, number>();
      groupRecords.forEach(r => {
        const nd = normalizeDob(r.dob);
        if (nd) dobCount.set(nd, (dobCount.get(nd) || 0) + 1);
      });
      let maxDCount = 0;
      let suggestedDob = distinctDobs[0];
      dobCount.forEach((count, d) => {
        if (count > maxDCount) {
          maxDCount = count;
          suggestedDob = d;
        }
      });

      const groupKey = `SAME_PASS_DIFF_DOB_${normPassportKey}_${distinctDobs.sort().join('_')}`;
      if (!processedGroupKeys.has(groupKey)) {
        processedGroupKeys.add(groupKey);
        suspectGroups.push({
          id: groupKey,
          type: 'SAME_PASSPORT_DIFF_DOB',
          title: `⚠️ Passport တူသော်လည်း မွေးသက္ကရာဇ် (DOB) ကွဲလွဲနေပါသည် (${distinctPList[0]})`,
          description: `Passport နံပါတ် '${distinctPList[0]}' တူညီနေသော်လည်း မွေးသက္ကရာဇ် (${distinctDobs.join(' vs ')}) ကွဲလွဲနေပါသည်။ Data Entry စာရိုက်မှားယွင်းမှု သို့မဟုတ် မတူညီသော ခရီးသည်မှတ်တမ်း ရောထွေးနေခြင်း ဖြစ်နိုင်ပါသည်။`,
          badgeColor: 'bg-red-600 text-white',
          records: groupRecords,
          suggestedPassport: distinctPList[0],
          suggestedName: distinctNList[0],
          suggestedDob: suggestedDob,
          distinctPassports: distinctPList,
          distinctNames: distinctNList,
          distinctDobs: distinctDobs,
          distinctNationalities: distinctNatList,
          confidenceScore: 90,
          reasonNote: `DOB Mismatch on Passport ${distinctPList[0]}`
        });
      }
    }
  });

  // --- CATEGORY 3: SAME / SIMILAR NAME BUT DIFFERENT PASSPORTS (WITHOUT EXACT DOB MATCH) ---
  const nameToRecordsMap = new Map<string, ImmRecord[]>();
  for (const r of validRecords) {
    const key = tokenSortedName(r.fullname);
    if (!key || key.length < 3) continue;
    if (!nameToRecordsMap.has(key)) {
      nameToRecordsMap.set(key, []);
    }
    nameToRecordsMap.get(key)!.push(r);
  }

  nameToRecordsMap.forEach((groupRecords, sortedNameKey) => {
    const passportMap = new Map<string, number>();
    groupRecords.forEach(r => {
      const np = normalizePassport(r.passport);
      passportMap.set(np, (passportMap.get(np) || 0) + 1);
    });

    if (passportMap.size > 1) {
      let maxCount = 0;
      let suggestedP = '';
      passportMap.forEach((count, p) => {
        if (count > maxCount) {
          maxCount = count;
          suggestedP = p;
        }
      });
      const rawSuggestedP = groupRecords.find(r => normalizePassport(r.passport) === suggestedP)?.passport || suggestedP;
      const rawSuggestedName = groupRecords.find(r => r.fullname)?.fullname || sortedNameKey;

      const distinctPList = Array.from(new Set(groupRecords.map(r => r.passport.toUpperCase().trim())));
      const distinctNList = Array.from(new Set(groupRecords.map(r => r.fullname.trim())));
      const distinctDList = Array.from(new Set(groupRecords.map(r => normalizeDob(r.dob)).filter(Boolean)));
      const distinctNatList = Array.from(new Set(groupRecords.map(r => r.nationality || '').filter(Boolean)));

      const isAlreadyChain = distinctPList.every(p => {
        const chain = linkedChainMap.get(normalizePassport(p));
        return chain && distinctPList.every(other => chain.includes(normalizePassport(other)));
      });

      // Avoid duplicating groups that were already flagged as EXACT SAME_NAME_DOB_DIFF_PASSPORT
      const exactDobGroupKey = `SAME_NAME_DOB_${sortedNameKey}__DOB__${distinctDList[0] || ''}_${distinctPList.sort().join('_')}`;
      if (distinctDList.length === 1 && processedGroupKeys.has(exactDobGroupKey)) {
        return;
      }

      const groupKey = `SAME_NAME_${sortedNameKey}_${distinctPList.sort().join('_')}`;
      if (!processedGroupKeys.has(groupKey)) {
        processedGroupKeys.add(groupKey);
        suspectGroups.push({
          id: groupKey,
          type: isAlreadyChain ? 'DUAL_PASSPORT_LINKED' : 'SAME_NAME_DIFF_PASSPORT',
          title: isAlreadyChain 
            ? `🔗 အတည်ပြုပြီး Dual / Renewed Passports (${distinctPList.join(' ⇄ ')})`
            : `အမည်တူသော်လည်း Passport နံပါတ် မတူပါ (${rawSuggestedName})`,
          description: isAlreadyChain
            ? `ပုဂ္ဂိုလ် (${rawSuggestedName}) ၏ Passport နံပါတ်များ ချိတ်ဆက်ပြီး ဖြစ်ပါသည်။`
            : `အမည် '${rawSuggestedName}' အတွက် Passport နံပါတ် (${distinctPList.join(', ')}) ကွာခြားချက် ရှာဖွေတွေ့ရှိပါသည်။${distinctDList.length ? ` (DOB: ${distinctDList.join(', ')})` : ''} (Passport အသစ်လဲခြင်း/Dual Passports သို့မဟုတ် စာလုံးမှားယွင်းမှု ဖြစ်နိုင်ပါသည်)`,
          badgeColor: isAlreadyChain ? 'bg-emerald-600 text-white' : 'bg-rose-500 text-white',
          records: groupRecords,
          suggestedPassport: rawSuggestedP,
          suggestedName: rawSuggestedName,
          suggestedDob: distinctDList[0] || '',
          distinctPassports: distinctPList,
          distinctNames: distinctNList,
          distinctDobs: distinctDList,
          distinctNationalities: distinctNatList,
          isAlreadyLinked: isAlreadyChain,
          linkedPassports: distinctPList,
          confidenceScore: distinctDList.length > 0 ? 80 : 65
        });
      }
    }
  });

  // --- CATEGORY 4: SAME PASSPORT BUT DIFFERENT NAMES ---
  passportToRecordsMap.forEach((groupRecords, normPassportKey) => {
    const nameMap = new Map<string, number>();
    groupRecords.forEach(r => {
      const tn = tokenSortedName(r.fullname);
      nameMap.set(tn, (nameMap.get(tn) || 0) + 1);
    });

    if (nameMap.size > 1) {
      let maxCount = 0;
      let suggestedTN = '';
      nameMap.forEach((count, tn) => {
        if (count > maxCount) {
          maxCount = count;
          suggestedTN = tn;
        }
      });

      const rawSuggestedP = groupRecords.find(r => normalizePassport(r.passport) === normPassportKey)?.passport || normPassportKey;
      const rawSuggestedName = groupRecords.find(r => tokenSortedName(r.fullname) === suggestedTN)?.fullname || '';

      const distinctPList = Array.from(new Set(groupRecords.map(r => r.passport.toUpperCase().trim())));
      const distinctNList = Array.from(new Set(groupRecords.map(r => r.fullname.trim())));
      const distinctDList = Array.from(new Set(groupRecords.map(r => normalizeDob(r.dob)).filter(Boolean)));
      const distinctNatList = Array.from(new Set(groupRecords.map(r => r.nationality || '').filter(Boolean)));

      const groupKey = `SAME_PASS_${normPassportKey}_${distinctNList.sort().join('_')}`;
      if (!processedGroupKeys.has(groupKey)) {
        processedGroupKeys.add(groupKey);
        suspectGroups.push({
          id: groupKey,
          type: 'SAME_PASSPORT_DIFF_NAME',
          title: `Passport နံပါတ်တူသော်လည်း အမည် ကွာဟနေပါသည် (${rawSuggestedP})`,
          description: `Passport နံပါတ် '${rawSuggestedP}' တွင် အမည်များ (${distinctNList.join(', ')}) မတူညီဘဲ ဖြစ်နေပါသည်။`,
          badgeColor: 'bg-amber-500 text-white',
          records: groupRecords,
          suggestedPassport: rawSuggestedP,
          suggestedName: rawSuggestedName,
          suggestedDob: distinctDList[0] || '',
          distinctPassports: distinctPList,
          distinctNames: distinctNList,
          distinctDobs: distinctDList,
          distinctNationalities: distinctNatList,
          confidenceScore: 75
        });
      }
    }
  });

  // --- CATEGORY 5: SIMILAR PASSPORTS (TYPO / LEVENSHTEIN MATCH) ---
  const allDistinctPassports = Array.from(passportToRecordsMap.keys());
  for (let i = 0; i < allDistinctPassports.length; i++) {
    for (let j = i + 1; j < allDistinctPassports.length; j++) {
      const p1 = allDistinctPassports[i];
      const p2 = allDistinctPassports[j];

      if (isSimilarPassport(p1, p2)) {
        const recordsP1 = passportToRecordsMap.get(p1) || [];
        const recordsP2 = passportToRecordsMap.get(p2) || [];

        const namesP1 = recordsP1.map(r => tokenSortedName(r.fullname));
        const namesP2 = recordsP2.map(r => tokenSortedName(r.fullname));
        const dobsP1 = recordsP1.map(r => normalizeDob(r.dob)).filter(Boolean);
        const dobsP2 = recordsP2.map(r => normalizeDob(r.dob)).filter(Boolean);

        const nameMatch = namesP1.some(n1 => namesP2.some(n2 => n1 === n2 || (n1.length > 3 && n2.length > 3 && levenshteinDistance(n1, n2) <= 3)));
        const dobMatch = dobsP1.some(d1 => dobsP2.includes(d1));

        if (nameMatch || dobMatch) {
          const combinedRecords = [...recordsP1, ...recordsP2];
          const distinctPList = Array.from(new Set(combinedRecords.map(r => r.passport.toUpperCase().trim())));
          const distinctNList = Array.from(new Set(combinedRecords.map(r => r.fullname.trim())));
          const distinctDList = Array.from(new Set(combinedRecords.map(r => normalizeDob(r.dob)).filter(Boolean)));
          const distinctNatList = Array.from(new Set(combinedRecords.map(r => r.nationality || '').filter(Boolean)));

          const suggestedP = recordsP1.length >= recordsP2.length ? recordsP1[0].passport : recordsP2[0].passport;
          const suggestedN = combinedRecords[0].fullname;
          const suggestedDob = distinctDList[0] || '';

          const groupKey = `SIMILAR_PASS_${p1}_${p2}`;
          if (!processedGroupKeys.has(groupKey)) {
            processedGroupKeys.add(groupKey);
            suspectGroups.push({
              id: groupKey,
              type: 'SIMILAR_PASSPORT_TYPO',
              title: `ဆင်တူ Passport နံပါတ်များ (Typo/စာလုံးလွဲမှားမှု ဖြစ်နိုင်ပါသည်)`,
              description: `Passport နံပါတ် '${distinctPList.join("' နှင့် '")}' သည် စာလုံး ၁-၂ လုံးသာ ကွာခြားပြီး အမည်တူ/ဆင်တူသူများ${distinctDList.length ? ` (DOB: ${distinctDList.join(', ')})` : ''} ဖြစ်ပါသည်။`,
              badgeColor: 'bg-indigo-600 text-white',
              records: combinedRecords,
              suggestedPassport: suggestedP,
              suggestedName: suggestedN,
              suggestedDob: suggestedDob,
              distinctPassports: distinctPList,
              distinctNames: distinctNList,
              distinctDobs: distinctDList,
              distinctNationalities: distinctNatList,
              confidenceScore: dobMatch ? 95 : 80
            });
          }
        }
      }
    }
  }

  // Sort groups: Conflicting DOB errors & High Confidence DOB Matches first, then others
  return suspectGroups.sort((a, b) => {
    const priority = {
      SAME_PASSPORT_DIFF_DOB: 1,
      SAME_NAME_DOB_DIFF_PASSPORT: 2,
      SIMILAR_PASSPORT_TYPO: 3,
      SAME_NAME_DIFF_PASSPORT: 4,
      SAME_PASSPORT_DIFF_NAME: 5,
      DUAL_PASSPORT_LINKED: 6
    };
    return (priority[a.type] || 99) - (priority[b.type] || 99);
  });
}
