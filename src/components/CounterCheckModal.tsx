import React, { useState, useMemo } from 'react';
import { ImmRecord } from '../types';
import { analyzeRecordsForErrors, SuspectGroup, normalizePassport, normalizeDob } from '../utils/counterCheck';
import { 
  X, Search, AlertTriangle, CheckCircle2, Edit2, RefreshCw, 
  ShieldAlert, Sparkles, Check, Link2, Unlink, BookmarkCheck, 
  Calendar, User, Globe, FileText, CheckCircle, AlertCircle
} from 'lucide-react';

interface CounterCheckModalProps {
  isOpen: boolean;
  onClose: () => void;
  records: ImmRecord[];
  setRecords: React.Dispatch<React.SetStateAction<ImmRecord[]>>;
  showToast: (msg: string) => void;
  onEditRecord: (record: ImmRecord) => void;
}

export const CounterCheckModal: React.FC<CounterCheckModalProps> = ({
  isOpen,
  onClose,
  records,
  setRecords,
  showToast,
  onEditRecord
}) => {
  const [filterType, setFilterType] = useState<
    'ALL' | 'SAME_NAME_DOB_DIFF_PASSPORT' | 'SAME_PASSPORT_DIFF_DOB' | 'DUAL_PASSPORT_LINKED' | 'SAME_NAME_DIFF_PASSPORT' | 'SAME_PASSPORT_DIFF_NAME' | 'SIMILAR_PASSPORT_TYPO'
  >('ALL');
  const [searchQuery, setSearchQuery] = useState('');
  const [refreshTrigger, setRefreshTrigger] = useState(0);

  // Selected fixes per group ID: { [groupId]: { passport: string, name: string, dob?: string } }
  const [fixSelections, setFixSelections] = useState<Record<string, { passport: string; name: string; dob?: string }>>({});

  // Dual passport linking state per group ID: { [groupId]: { primaryPassport: string; secondaryPassport: string; remarks: string } }
  const [dualLinkState, setDualLinkState] = useState<Record<string, { primaryPassport: string; secondaryPassport: string; remarks: string }>>({});

  // Compute suspect groups dynamically from records with enhanced DOB algorithm
  const suspectGroups = useMemo(() => {
    return analyzeRecordsForErrors(records);
  }, [records, refreshTrigger]);

  // Filter groups
  const filteredGroups = useMemo(() => {
    return suspectGroups.filter(group => {
      if (filterType !== 'ALL' && group.type !== filterType) return false;
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const matchTitle = group.title.toLowerCase().includes(q);
        const matchDesc = group.description.toLowerCase().includes(q);
        const matchPassports = group.distinctPassports.some(p => p.toLowerCase().includes(q));
        const matchNames = group.distinctNames.some(n => n.toLowerCase().includes(q));
        const matchDobs = group.distinctDobs.some(d => d.toLowerCase().includes(q));
        return matchTitle || matchDesc || matchPassports || matchNames || matchDobs;
      }
      return true;
    });
  }, [suspectGroups, filterType, searchQuery]);

  if (!isOpen) return null;

  // Handler to bulk unify passport, name, and optionally DOB for a suspect group (Fix Typo)
  const handleUnifyGroup = (group: SuspectGroup) => {
    const selectedFix = fixSelections[group.id] || {
      passport: group.suggestedPassport,
      name: group.suggestedName,
      dob: group.suggestedDob || ''
    };

    const targetPassport = selectedFix.passport.trim().toUpperCase();
    const targetName = selectedFix.name.trim();
    const targetDob = selectedFix.dob?.trim() || '';

    if (!targetPassport) {
      showToast('⚠️ Passport နံပါတ် အလွတ်ဖြစ်နေပါသည်');
      return;
    }

    const recordIdsToUpdate = new Set(group.records.map(r => r.id));
    const nowIso = new Date().toISOString();

    setRecords(prev => {
      const updated = prev.map(r => {
        if (recordIdsToUpdate.has(r.id)) {
          return {
            ...r,
            passport: targetPassport,
            fullname: targetName || r.fullname,
            ...(targetDob ? { dob: targetDob } : {}),
            syncStatus: 'pending_sync' as const,
            updatedAt: nowIso
          };
        }
        return r;
      });

      try {
        localStorage.setItem('imm_records_react', JSON.stringify(updated));
      } catch (e) {}

      return updated;
    });

    showToast(`✅ မှတ်တမ်း ${group.records.length} ခု၏ Passport အား '${targetPassport}' ${targetDob ? `(DOB: ${targetDob})` : ''} သို့ ပေါင်းစပ်ပြင်ဆင်ပြီးပါပြီ`);
    setRefreshTrigger(prev => prev + 1);
  };

  // Handler to Link Passports as Dual / Renewed Passports (Unified Foreigner Profile)
  const handleLinkDualPassports = (group: SuspectGroup) => {
    const defaultPrimary = group.distinctPassports[0] || group.suggestedPassport;
    const defaultSecondary = group.distinctPassports.find(p => p !== defaultPrimary) || group.distinctPassports[1] || '';

    const currentDual = dualLinkState[group.id] || {
      primaryPassport: defaultPrimary,
      secondaryPassport: defaultSecondary,
      remarks: group.dualPassportRemarks || (group.suggestedDob ? `Dual / Renewed Passport (DOB: ${group.suggestedDob})` : 'Dual / Renewed Passport Record')
    };

    const primaryP = (currentDual.primaryPassport || defaultPrimary).toUpperCase().trim();
    const secondaryP = (currentDual.secondaryPassport || defaultSecondary).toUpperCase().trim();
    const dualRemark = currentDual.remarks || 'Dual / Renewed Passport Record';

    if (!primaryP || !secondaryP || primaryP === secondaryP) {
      showToast('⚠️ မတူညီသော Passport နံပါတ် (၂) ခု ရွေးချယ်ပေးပါ');
      return;
    }

    const allGroupPassports = Array.from(new Set([primaryP, secondaryP, ...group.distinctPassports]));
    const recordIdsToUpdate = new Set(group.records.map(r => r.id));
    const nowIso = new Date().toISOString();

    setRecords(prev => {
      const updated = prev.map(r => {
        const rNormP = normalizePassport(r.passport);
        const inGroup = recordIdsToUpdate.has(r.id) || allGroupPassports.some(p => normalizePassport(p) === rNormP);
        
        if (inGroup) {
          const isPrimary = rNormP === normalizePassport(primaryP);
          return {
            ...r,
            previousPassport: isPrimary ? secondaryP : (r.previousPassport || secondaryP),
            linkedPassports: allGroupPassports,
            dualPassportRemarks: dualRemark,
            syncStatus: 'pending_sync' as const,
            updatedAt: nowIso
          };
        }
        return r;
      });

      try {
        localStorage.setItem('imm_records_react', JSON.stringify(updated));
      } catch (e) {}

      return updated;
    });

    showToast(`🔗 Passport [${primaryP}] နှင့် [${secondaryP}] အား Dual / Renewed Profile အဖြစ် ချိတ်ဆက်မှတ်တမ်းတင်ပြီးပါပြီ`);
    setRefreshTrigger(prev => prev + 1);
  };

  // Handler to Unlink Passports
  const handleUnlinkPassports = (group: SuspectGroup) => {
    const recordIdsToUpdate = new Set(group.records.map(r => r.id));
    const nowIso = new Date().toISOString();

    setRecords(prev => {
      const updated = prev.map(r => {
        if (recordIdsToUpdate.has(r.id)) {
          return {
            ...r,
            previousPassport: undefined,
            linkedPassports: undefined,
            dualPassportRemarks: undefined,
            syncStatus: 'pending_sync' as const,
            updatedAt: nowIso
          };
        }
        return r;
      });

      try {
        localStorage.setItem('imm_records_react', JSON.stringify(updated));
      } catch (e) {}

      return updated;
    });

    showToast(`⛓️ Passport ချိတ်ဆက်မှုအား ဖြုတ်သိမ်းပြီးပါပြီ`);
    setRefreshTrigger(prev => prev + 1);
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/80 backdrop-blur-sm flex items-center justify-center p-2 sm:p-4 overflow-y-auto no-print font-pyidaungsu">
      <div className="bg-white border border-slate-200 rounded-3xl shadow-2xl w-full max-w-7xl max-h-[94vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-200">
        
        {/* HEADER */}
        <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-blue-950 text-white p-4 sm:p-5 flex items-center justify-between shrink-0 border-b border-white/10">
          <div className="flex items-center gap-3">
            <div className="bg-indigo-500/20 p-2.5 rounded-2xl border border-indigo-400/30 text-indigo-300 shrink-0 shadow-inner">
              <ShieldAlert size={26} />
            </div>
            <div>
              <div className="flex items-center gap-2 flex-wrap">
                <h2 className="text-base sm:text-lg font-black uppercase tracking-tight text-white">
                  Flight History Counter Check & Smart FFE Intelligence
                </h2>
                <span className="bg-rose-500/30 text-rose-300 border border-rose-400/40 text-[10px] font-black px-2.5 py-0.5 rounded-full uppercase font-mono">
                  {suspectGroups.length} Flagged
                </span>
                <span className="bg-emerald-500/20 text-emerald-300 border border-emerald-400/30 text-[10px] font-bold px-2 py-0.5 rounded-full flex items-center gap-1">
                  <Calendar size={11} /> DOB Multi-Verification Active
                </span>
              </div>
              <p className="text-[11px] text-slate-300 font-medium mt-0.5">
                လက်ရှိ Date of Birth (DOB) နှင့် ပြည့်စုံသော FFE အချက်အလက်များဖြင့် Passport အသစ်လဲခြင်း (Dual/Renewed Passports) နှင့် စာလုံးမှားယွင်းမှုများကို အဆင့်မြင့် ခွဲခြမ်းစိစစ်ခြင်း
              </p>
            </div>
          </div>
          
          <button
            onClick={onClose}
            className="p-2 text-slate-400 hover:text-white hover:bg-white/10 rounded-2xl transition-all cursor-pointer"
            title="Close"
          >
            <X size={20} />
          </button>
        </div>

        {/* CONTROLS & FILTERS */}
        <div className="bg-slate-50 border-b border-slate-200 p-3 sm:p-4 flex flex-col xl:flex-row items-stretch xl:items-center justify-between gap-3 shrink-0">
          
          {/* Filter Pills */}
          <div className="flex flex-wrap items-center gap-1.5 text-xs font-bold">
            <button
              onClick={() => setFilterType('ALL')}
              className={`px-3 py-1.5 rounded-xl transition-all flex items-center gap-1.5 cursor-pointer text-xs ${
                filterType === 'ALL'
                  ? 'bg-indigo-600 text-white shadow-md'
                  : 'bg-white text-slate-600 border border-slate-200 hover:bg-slate-100'
              }`}
            >
              <span>အားလုံး ({suspectGroups.length})</span>
            </button>

            {/* High-confidence Name + DOB Match */}
            <button
              onClick={() => setFilterType('SAME_NAME_DOB_DIFF_PASSPORT')}
              className={`px-3 py-1.5 rounded-xl transition-all flex items-center gap-1.5 cursor-pointer text-xs ${
                filterType === 'SAME_NAME_DOB_DIFF_PASSPORT'
                  ? 'bg-purple-600 text-white shadow-md'
                  : 'bg-white text-purple-700 border border-purple-200 hover:bg-purple-50'
              }`}
            >
              <Sparkles size={13} className="text-amber-300" />
              <span>အမည် + DOB တူ/PP မတူ ({suspectGroups.filter(g => g.type === 'SAME_NAME_DOB_DIFF_PASSPORT').length})</span>
            </button>

            {/* Conflicting DOB error */}
            <button
              onClick={() => setFilterType('SAME_PASSPORT_DIFF_DOB')}
              className={`px-3 py-1.5 rounded-xl transition-all flex items-center gap-1.5 cursor-pointer text-xs ${
                filterType === 'SAME_PASSPORT_DIFF_DOB'
                  ? 'bg-red-600 text-white shadow-md'
                  : 'bg-white text-red-700 border border-red-200 hover:bg-red-50'
              }`}
            >
              <AlertCircle size={13} />
              <span>PPတူ/DOBလွဲမှား ({suspectGroups.filter(g => g.type === 'SAME_PASSPORT_DIFF_DOB').length})</span>
            </button>

            <button
              onClick={() => setFilterType('DUAL_PASSPORT_LINKED')}
              className={`px-3 py-1.5 rounded-xl transition-all flex items-center gap-1.5 cursor-pointer text-xs ${
                filterType === 'DUAL_PASSPORT_LINKED'
                  ? 'bg-emerald-600 text-white shadow-md'
                  : 'bg-white text-emerald-700 border border-emerald-200 hover:bg-emerald-50'
              }`}
            >
              <Link2 size={13} />
              <span>အတည်ပြုပြီး Dual PP ({suspectGroups.filter(g => g.type === 'DUAL_PASSPORT_LINKED').length})</span>
            </button>

            <button
              onClick={() => setFilterType('SAME_NAME_DIFF_PASSPORT')}
              className={`px-3 py-1.5 rounded-xl transition-all flex items-center gap-1.5 cursor-pointer text-xs ${
                filterType === 'SAME_NAME_DIFF_PASSPORT'
                  ? 'bg-rose-600 text-white shadow-md'
                  : 'bg-white text-slate-600 border border-slate-200 hover:bg-slate-100'
              }`}
            >
              <AlertTriangle size={13} />
              <span>အမည်တူ/PPမတူ ({suspectGroups.filter(g => g.type === 'SAME_NAME_DIFF_PASSPORT').length})</span>
            </button>

            <button
              onClick={() => setFilterType('SAME_PASSPORT_DIFF_NAME')}
              className={`px-3 py-1.5 rounded-xl transition-all flex items-center gap-1.5 cursor-pointer text-xs ${
                filterType === 'SAME_PASSPORT_DIFF_NAME'
                  ? 'bg-amber-600 text-white shadow-md'
                  : 'bg-white text-slate-600 border border-slate-200 hover:bg-slate-100'
              }`}
            >
              <span>PPတူ/အမည်မတူ ({suspectGroups.filter(g => g.type === 'SAME_PASSPORT_DIFF_NAME').length})</span>
            </button>

            <button
              onClick={() => setFilterType('SIMILAR_PASSPORT_TYPO')}
              className={`px-3 py-1.5 rounded-xl transition-all flex items-center gap-1.5 cursor-pointer text-xs ${
                filterType === 'SIMILAR_PASSPORT_TYPO'
                  ? 'bg-slate-800 text-white shadow-md'
                  : 'bg-white text-slate-600 border border-slate-200 hover:bg-slate-100'
              }`}
            >
              <span>ဆင်တူ Typo ({suspectGroups.filter(g => g.type === 'SIMILAR_PASSPORT_TYPO').length})</span>
            </button>
          </div>

          {/* Search Input & Refresh */}
          <div className="flex items-center gap-2">
            <div className="relative flex-1 sm:w-72">
              <input
                type="text"
                placeholder="Search Passport, Name or DOB..."
                value={searchQuery}
                onChange={e => setSearchQuery(e.target.value)}
                className="w-full pl-8 pr-3 py-1.5 text-xs font-bold border border-slate-200 rounded-xl bg-white outline-none focus:ring-2 focus:ring-indigo-500 shadow-2xs"
              />
              <Search size={13} className="absolute left-2.5 top-2.5 text-slate-400" />
            </div>
            
            <button
              onClick={() => setRefreshTrigger(prev => prev + 1)}
              className="p-2 text-slate-600 bg-white border border-slate-200 hover:bg-indigo-50 hover:text-indigo-600 rounded-xl transition-all cursor-pointer shrink-0 shadow-2xs"
              title="Re-run analysis"
            >
              <RefreshCw size={14} />
            </button>
          </div>
        </div>

        {/* CONTENT BODY */}
        <div className="flex-1 overflow-y-auto p-3 sm:p-5 space-y-4 bg-slate-100/70">
          
          {filteredGroups.length === 0 ? (
            <div className="bg-white border border-emerald-200 rounded-3xl p-8 sm:p-12 text-center shadow-sm my-6 space-y-3">
              <div className="w-16 h-16 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center mx-auto shadow-inner">
                <CheckCircle2 size={36} />
              </div>
              <h3 className="text-base sm:text-lg font-black text-slate-800">
                {suspectGroups.length === 0
                  ? 'သံသယဖြစ်ဖွယ် အမည်၊ မွေးသက္ကရာဇ် (DOB) နှင့် Passport စာရိုက်မှားယွင်းမှု မရှိပါ။'
                  : 'ရွေးချယ်ထားသော စိစစ်မှုနှင့် ကိုက်ညီသည့် မှတ်တမ်း မရှိပါ။'}
              </h3>
              <p className="text-xs text-slate-500 max-w-lg mx-auto">
                {suspectGroups.length === 0
                  ? 'လက်ရှိ Flight History နှင့် FFE မှတ်တမ်းများအားလုံးတွင် အမည်၊ DOB နှင့် Passport နံပါတ်များ တညီတညွတ်တည်း မှန်ကန်စွာ ရှိနေပါသည်။'
                  : 'အခြား စိစစ်မှု အမျိုးအစား သို့မဟုတ် ရှာဖွေစကားလုံး ပြောင်းလဲစစ်ဆေးပါ။'}
              </p>
            </div>
          ) : (
            filteredGroups.map(group => {
              const currentFix = fixSelections[group.id] || {
                passport: group.suggestedPassport,
                name: group.suggestedName,
                dob: group.suggestedDob || ''
              };

              const defaultPrimary = group.distinctPassports[0] || group.suggestedPassport;
              const defaultSecondary = group.distinctPassports.find(p => p !== defaultPrimary) || group.distinctPassports[1] || '';

              const currentDual = dualLinkState[group.id] || {
                primaryPassport: defaultPrimary,
                secondaryPassport: defaultSecondary,
                remarks: group.dualPassportRemarks || (group.suggestedDob ? `Dual / Renewed Passport (DOB: ${group.suggestedDob})` : 'Dual / Renewed Passport Record')
              };

              const isLinkedGroup = group.isAlreadyLinked || group.type === 'DUAL_PASSPORT_LINKED';
              const isHighConfidenceNameDob = group.type === 'SAME_NAME_DOB_DIFF_PASSPORT';
              const isConflictingDob = group.type === 'SAME_PASSPORT_DIFF_DOB';

              return (
                <div
                  key={group.id}
                  className={`bg-white border ${
                    isLinkedGroup 
                      ? 'border-emerald-300 ring-1 ring-emerald-200' 
                      : isConflictingDob
                        ? 'border-red-300 ring-1 ring-red-200'
                        : isHighConfidenceNameDob
                          ? 'border-purple-300 ring-1 ring-purple-200'
                          : 'border-slate-200'
                  } rounded-2xl shadow-sm overflow-hidden transition-all hover:shadow-md`}
                >
                  {/* GROUP HEADER */}
                  <div className={`p-3.5 sm:p-4 ${
                    isLinkedGroup 
                      ? 'bg-emerald-50/80 border-b border-emerald-200' 
                      : isConflictingDob
                        ? 'bg-red-50/80 border-b border-red-200'
                        : isHighConfidenceNameDob
                          ? 'bg-purple-50/80 border-b border-purple-200'
                          : 'bg-slate-50 border-b border-slate-200'
                  } flex flex-col sm:flex-row sm:items-center justify-between gap-2`}>
                    <div className="space-y-1">
                      <div className="flex items-center gap-2 flex-wrap">
                        <span className={`text-[10px] font-black px-2 py-0.5 rounded-md uppercase font-mono ${group.badgeColor}`}>
                          {group.type}
                        </span>
                        
                        {group.confidenceScore && (
                          <span className="text-[10px] font-black bg-white/90 border border-slate-200 px-2 py-0.5 rounded-md text-slate-700 font-mono">
                            Match Score: {group.confidenceScore}%
                          </span>
                        )}

                        <h4 className="text-xs sm:text-sm font-black text-slate-900 flex items-center gap-2">
                          {group.title}
                          {isLinkedGroup && (
                            <span className="bg-emerald-600 text-white text-[10px] px-2 py-0.5 rounded-full font-bold">
                              Unified Chain Active
                            </span>
                          )}
                        </h4>
                      </div>

                      <p className="text-[11px] text-slate-700 font-medium">
                        {group.description}
                      </p>

                      {group.dualPassportRemarks && (
                        <div className="text-[11px] text-emerald-800 font-bold bg-emerald-100/90 px-2.5 py-1 rounded-lg border border-emerald-300 inline-block">
                          📝 မှတ်ချက်: {group.dualPassportRemarks}
                        </div>
                      )}
                    </div>

                    <div className="flex items-center gap-2 shrink-0">
                      <div className="text-[10px] font-bold text-slate-600 bg-white border border-slate-200 px-2.5 py-1 rounded-lg shadow-2xs">
                        သက်ဆိုင်ရာ မှတ်တမ်း: <span className="font-mono font-black text-slate-900">{group.records.length}</span> ခု
                      </div>
                    </div>
                  </div>

                  {/* RECORDS TABLE WITH COMPLETE FFE & DOB DATA */}
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse text-xs">
                      <thead>
                        <tr className="bg-slate-100 text-slate-600 text-[10px] uppercase tracking-wider font-black border-b border-slate-200">
                          <th className="p-2.5">ရက်စွဲ/အချိန်</th>
                          <th className="p-2.5">ယာဉ်/လေယာဉ်</th>
                          <th className="p-2.5 text-center">In/Out</th>
                          <th className="p-2.5">အမည် (Full Name)</th>
                          <th className="p-2.5">မွေးသက္ကရာဇ် (DOB)</th>
                          <th className="p-2.5">Passport နံပါတ်</th>
                          <th className="p-2.5">Previous/Dual PP</th>
                          <th className="p-2.5">နိုင်ငံသား</th>
                          <th className="p-2.5">ဗီဇာ / အချက်အလက်</th>
                          <th className="p-2.5 text-center">Action</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 font-medium text-slate-800">
                        {group.records.map((r, idx) => {
                          const isSuggestedP = normalizePassport(r.passport) === normalizePassport(group.suggestedPassport);
                          const isSuggestedD = group.suggestedDob && normalizeDob(r.dob) === normalizeDob(group.suggestedDob);

                          return (
                            <tr key={r.id || idx} className="hover:bg-indigo-50/40 transition-colors">
                              <td className="p-2.5 font-mono text-[11px] text-slate-600 whitespace-nowrap">{r.timestamp || '-'}</td>
                              <td className="p-2.5 font-bold text-slate-900 whitespace-nowrap">{r.vehicleInfo || '-'}</td>
                              <td className="p-2.5 text-center">
                                <span
                                  className={`px-2 py-0.5 rounded-full text-[9px] font-black uppercase ${
                                    r.mode === 'IN'
                                      ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                                      : 'bg-rose-100 text-rose-800 border border-rose-300'
                                  }`}
                                >
                                  {r.mode}
                                </span>
                              </td>
                              <td className="p-2.5 font-bold text-slate-900 whitespace-nowrap">{r.fullname}</td>
                              <td className="p-2.5 font-mono font-bold whitespace-nowrap">
                                {r.dob ? (
                                  <span className={`px-2 py-0.5 rounded-md text-[11px] flex items-center gap-1 w-fit ${
                                    isConflictingDob
                                      ? 'bg-red-100 text-red-900 border border-red-300'
                                      : 'bg-purple-50 text-purple-900 border border-purple-200'
                                  }`}>
                                    <Calendar size={11} className="text-purple-600" />
                                    {r.dob}
                                  </span>
                                ) : (
                                  <span className="text-slate-400 italic text-[11px]">No DOB</span>
                                )}
                              </td>
                              <td className="p-2.5 font-mono font-black whitespace-nowrap">
                                <span
                                  className={`px-2 py-0.5 rounded-md ${
                                    isSuggestedP
                                      ? 'bg-emerald-50 text-emerald-800 border border-emerald-300'
                                      : 'bg-amber-100 text-amber-900 border border-amber-300 font-extrabold'
                                  }`}
                                >
                                  {r.passport}
                                </span>
                              </td>
                              <td className="p-2.5 font-mono text-[11px] whitespace-nowrap">
                                {r.previousPassport ? (
                                  <span className="bg-purple-100 text-purple-900 border border-purple-300 px-2 py-0.5 rounded-md font-bold text-[10px]">
                                    🔗 {r.previousPassport}
                                  </span>
                                ) : (
                                  <span className="text-slate-400">-</span>
                                )}
                              </td>
                              <td className="p-2.5 text-slate-600 whitespace-nowrap">{r.nationality || '-'}</td>
                              <td className="p-2.5 text-[11px] text-slate-600 whitespace-nowrap">
                                <span className="font-semibold text-slate-800">{r.visaType || '-'}</span>
                                {r.visaNumber && <span className="text-slate-400 font-mono ml-1">({r.visaNumber})</span>}
                              </td>
                              <td className="p-2.5 text-center">
                                <button
                                  onClick={() => {
                                    onClose();
                                    onEditRecord(r);
                                  }}
                                  className="px-2.5 py-1 text-[10px] font-bold bg-slate-100 text-slate-700 hover:bg-blue-600 hover:text-white rounded-lg transition-all flex items-center gap-1 mx-auto cursor-pointer"
                                  title="Edit entry manually"
                                >
                                  <Edit2 size={11} />
                                  <span>Edit</span>
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>

                  {/* ACTION CONTROLS */}
                  <div className="p-3.5 bg-slate-50 border-t border-slate-200 flex flex-col gap-3">
                    
                    {/* OPTION 1: LINK AS DUAL / RENEWED PASSPORTS */}
                    <div className="p-3 bg-emerald-50/90 border border-emerald-200 rounded-xl flex flex-col lg:flex-row items-stretch lg:items-center justify-between gap-3">
                      <div className="flex flex-col sm:flex-row items-start sm:items-center gap-2 flex-1">
                        <div className="flex items-center gap-1.5 text-xs font-black text-emerald-950 shrink-0">
                          <Link2 size={15} className="text-emerald-700" />
                          <span>🔗 Dual / Renewed Passports အဖြစ် ချိတ်ဆက်ရန်:</span>
                        </div>

                        <div className="flex items-center gap-2 flex-wrap w-full sm:w-auto">
                          <div className="flex items-center gap-1 text-[11px] font-bold text-emerald-900">
                            <span>အဓိက/လက်ရှိ PP:</span>
                            <select
                              value={currentDual.primaryPassport}
                              onChange={e =>
                                setDualLinkState(prev => ({
                                  ...prev,
                                  [group.id]: { ...currentDual, primaryPassport: e.target.value }
                                }))
                              }
                              className="bg-white border border-emerald-300 text-xs font-mono font-black text-emerald-950 rounded-lg px-2 py-1 outline-none focus:ring-2 focus:ring-emerald-500"
                            >
                              {group.distinctPassports.map(p => (
                                <option key={p} value={p}>
                                  {p}
                                </option>
                              ))}
                            </select>
                          </div>

                          <div className="flex items-center gap-1 text-[11px] font-bold text-emerald-900">
                            <span>ယခင်/အဟောင်း PP:</span>
                            <select
                              value={currentDual.secondaryPassport}
                              onChange={e =>
                                setDualLinkState(prev => ({
                                  ...prev,
                                  [group.id]: { ...currentDual, secondaryPassport: e.target.value }
                                }))
                              }
                              className="bg-white border border-emerald-300 text-xs font-mono font-black text-emerald-950 rounded-lg px-2 py-1 outline-none focus:ring-2 focus:ring-emerald-500"
                            >
                              {group.distinctPassports.map(p => (
                                <option key={p} value={p}>
                                  {p}
                                </option>
                              ))}
                            </select>
                          </div>

                          <input
                            type="text"
                            placeholder="Passport နံပါတ် (၂) ခု ကိုင်ဆောင်သူ မှတ်ချက် (Remark)..."
                            value={currentDual.remarks}
                            onChange={e =>
                              setDualLinkState(prev => ({
                                ...prev,
                                [group.id]: { ...currentDual, remarks: e.target.value }
                              }))
                            }
                            className="flex-1 min-w-[200px] bg-white border border-emerald-300 text-xs text-emerald-950 rounded-lg px-2.5 py-1 outline-none focus:ring-2 focus:ring-emerald-500"
                          />
                        </div>
                      </div>

                      <div className="flex items-center gap-2 shrink-0">
                        <button
                          onClick={() => handleLinkDualPassports(group)}
                          className="px-4 py-1.5 text-xs font-black bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl shadow-sm hover:shadow-md transition-all flex items-center justify-center gap-1.5 cursor-pointer"
                        >
                          <BookmarkCheck size={14} />
                          <span>Dual / Renewed Passport အဖြစ် သိမ်းမည်</span>
                        </button>

                        {isLinkedGroup && (
                          <button
                            onClick={() => handleUnlinkPassports(group)}
                            className="px-3 py-1.5 text-xs font-bold bg-white text-rose-700 border border-rose-300 hover:bg-rose-50 rounded-xl transition-all flex items-center justify-center gap-1 cursor-pointer"
                            title="Unlink"
                          >
                            <Unlink size={13} />
                            <span>ချိတ်ဆက်မှုဖြုတ်မည်</span>
                          </button>
                        )}
                      </div>
                    </div>

                    {/* OPTION 2: QUICK UNIFY (FOR TYPO / DATA CORRECTION) */}
                    <div className="p-3 bg-indigo-50/70 border border-indigo-100 rounded-xl flex flex-col lg:flex-row items-stretch lg:items-center justify-between gap-3">
                      <div className="flex flex-col sm:flex-row items-start sm:items-center gap-2 flex-1">
                        <div className="text-[11px] font-bold text-indigo-950 shrink-0">
                          ⚡ Typo / စာလုံးမှား ပြင်ဆင်ရန် (Quick Unify):
                        </div>
                        
                        <div className="flex items-center gap-2 flex-wrap w-full sm:w-auto">
                          {/* Passport Selector */}
                          <select
                            value={currentFix.passport}
                            onChange={e =>
                              setFixSelections(prev => ({
                                ...prev,
                                [group.id]: { ...currentFix, passport: e.target.value }
                              }))
                            }
                            className="bg-white border border-indigo-200 text-xs font-mono font-black text-indigo-900 rounded-lg px-2 py-1 outline-none focus:ring-2 focus:ring-indigo-500 shadow-2xs"
                          >
                            {group.distinctPassports.map(p => (
                              <option key={p} value={p}>
                                PP: {p} {p === group.suggestedPassport ? '(Suggested)' : ''}
                              </option>
                            ))}
                          </select>

                          {/* Name Selector */}
                          <select
                            value={currentFix.name}
                            onChange={e =>
                              setFixSelections(prev => ({
                                ...prev,
                                [group.id]: { ...currentFix, name: e.target.value }
                              }))
                            }
                            className="bg-white border border-indigo-200 text-xs font-bold text-indigo-900 rounded-lg px-2 py-1 outline-none focus:ring-2 focus:ring-indigo-500 shadow-2xs"
                          >
                            {group.distinctNames.map(n => (
                              <option key={n} value={n}>
                                Name: {n}
                              </option>
                            ))}
                          </select>

                          {/* DOB Selector if available */}
                          {group.distinctDobs.length > 0 && (
                            <select
                              value={currentFix.dob}
                              onChange={e =>
                                setFixSelections(prev => ({
                                  ...prev,
                                  [group.id]: { ...currentFix, dob: e.target.value }
                                }))
                              }
                              className="bg-white border border-indigo-200 text-xs font-mono font-bold text-indigo-900 rounded-lg px-2 py-1 outline-none focus:ring-2 focus:ring-indigo-500 shadow-2xs"
                            >
                              {group.distinctDobs.map(d => (
                                <option key={d} value={d}>
                                  DOB: {d} {d === group.suggestedDob ? '(Suggested)' : ''}
                                </option>
                              ))}
                            </select>
                          )}
                        </div>
                      </div>

                      <button
                        onClick={() => handleUnifyGroup(group)}
                        className="px-4 py-1.5 text-xs font-black bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl shadow-sm hover:shadow-md transition-all flex items-center justify-center gap-1.5 cursor-pointer shrink-0"
                      >
                        <Check size={14} strokeWidth={3} />
                        <span>စာရိုက်မှားယွင်းမှု အားလုံး ({currentFix.passport}) သို့ ပြင်မည်</span>
                      </button>
                    </div>

                  </div>

                </div>
              );
            })
          )}

        </div>

        {/* FOOTER */}
        <div className="bg-slate-50 border-t border-slate-200 p-3 sm:p-4 flex items-center justify-between text-xs shrink-0">
          <span className="text-[11px] font-bold text-slate-500">
            * စနစ်အသစ်သည် Date of Birth (DOB) အချက်အလက်များဖြင့် တိကျစွာ ခွဲခြားစစ်ဆေးပေးပြီး Dual Passport ကိုင်ဆောင်သူများ၏ ခရီးသွားမှတ်တမ်းအားလုံးကို ချိတ်ဆက်ထိန်းသိမ်းပေးပါသည်။
          </span>
          <button
            onClick={onClose}
            className="px-5 py-2 font-black text-slate-700 bg-white border border-slate-300 hover:bg-slate-100 rounded-xl transition-all cursor-pointer shadow-xs"
          >
            Close
          </button>
        </div>

      </div>
    </div>
  );
};
