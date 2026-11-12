import React from 'react';
import { GitMerge, ArrowRightLeft, ShieldAlert, Check, X, Smartphone, Server, Layers } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { ImmRecord } from '../types';

interface ConflictResolverModalProps {
  isOpen: boolean;
  onClose: () => void;
  localRecord: ImmRecord;
  remoteRecord: ImmRecord;
  onResolve: (action: 'KEEP_LOCAL' | 'KEEP_REMOTE' | 'KEEP_BOTH') => void;
}

export const ConflictResolverModal: React.FC<ConflictResolverModalProps> = ({
  isOpen,
  onClose,
  localRecord,
  remoteRecord,
  onResolve,
}) => {
  if (!isOpen || !localRecord || !remoteRecord) return null;

  const compareFields: Array<{ label: string; key: keyof ImmRecord }> = [
    { label: 'Full Name (အမည်)', key: 'fullname' },
    { label: 'Passport No (ပတ်စပို့)', key: 'passport' },
    { label: 'Nationality (နိုင်ငံသား)', key: 'nationality' },
    { label: 'Gender (ကျား/မ)', key: 'gender' },
    { label: 'Visa Type (ဗီဇာ)', key: 'visaType' },
    { label: 'Visa Number (ဗီဇာအမှတ်)', key: 'visaNumber' },
    { label: 'Stay Location (လိပ်စာ)', key: 'address' },
    { label: 'Stay Description', key: 'stayDescription' },
    { label: 'Purpose of Visit', key: 'purposeOfVisit' },
    { label: 'Stay Dates', key: 'stayFrom' },
    { label: 'Vehicle Info (ကား/လေယာဉ်)', key: 'vehicleInfo' },
    { label: 'Agent / Brought By', key: 'broughtBy' },
    { label: 'Official Name (စာရင်းသွင်းအရာရှိ)', key: 'officialName' },
    { label: 'Updated Timestamp', key: 'updatedAt' },
  ];

  return (
    <AnimatePresence>
      <div className="fixed inset-0 z-[300] flex items-center justify-center p-4 select-text no-print">
        <motion.div
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          onClick={onClose}
          className="absolute inset-0 bg-slate-950/80 backdrop-blur-md"
        />

        <motion.div
          initial={{ scale: 0.9, opacity: 0, y: 20 }}
          animate={{ scale: 1, opacity: 1, y: 0 }}
          exit={{ scale: 0.9, opacity: 0, y: 20 }}
          className="relative bg-white dark:bg-slate-900 w-full max-w-4xl rounded-3xl overflow-hidden shadow-2xl border-2 border-amber-500/40 text-slate-900 dark:text-white max-h-[90vh] flex flex-col"
        >
          {/* Header */}
          <div className="bg-amber-950 p-6 border-b border-amber-800/60 flex items-start justify-between gap-4 text-white shrink-0">
            <div className="flex items-center gap-3">
              <div className="w-12 h-12 rounded-2xl bg-amber-900/80 text-amber-200 flex items-center justify-center shrink-0 border border-amber-700/50 shadow-inner animate-bounce">
                <GitMerge size={26} />
              </div>
              <div>
                <span className="text-[10px] font-black uppercase tracking-widest text-amber-300 block">
                  Multi-Device Data Sync Conflict Resolver
                </span>
                <h3 className="text-lg font-black text-white leading-tight mt-0.5">
                  ဒေတာတိုက်ဆိုင်မှု ကွဲလွဲချက် စစ်ဆေးရန် (Conflict Merge)
                </h3>
              </div>
            </div>
            <button
              onClick={onClose}
              className="p-1.5 rounded-full hover:bg-amber-900/60 text-amber-300 hover:text-white transition-colors cursor-pointer"
            >
              <X size={18} />
            </button>
          </div>

          <div className="p-6 overflow-y-auto space-y-6 flex-1">
            <div className="bg-amber-50 dark:bg-amber-950/40 border border-amber-200 dark:border-amber-800/80 p-4 rounded-2xl flex items-start gap-3">
              <ShieldAlert size={20} className="text-amber-600 dark:text-amber-400 shrink-0 mt-0.5" />
              <div className="text-xs font-semibold text-amber-950 dark:text-amber-200 leading-relaxed">
                စက်တစ်ခုထက်ပို၍ အချက်အလက်များ တိုင်ပြိုင် ပြင်ဆင်ထားသဖြင့် ဒေတာ မပျောက်ပျက်စေရန်အတွက် အောက်ပါ ကွဲလွဲချက်များကို တိုက်ဆိုင်စစ်ဆေးပြီး နှစ်သက်ရာ နည်းလမ်းကို ရွေးချယ်ပေးပါ-
              </div>
            </div>

            {/* Side by side comparison */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {/* Local Record Card */}
              <div className="bg-indigo-50/50 dark:bg-indigo-950/30 border-2 border-indigo-200 dark:border-indigo-800 rounded-2xl p-4 space-y-3">
                <div className="flex items-center justify-between border-b border-indigo-200 dark:border-indigo-800 pb-2">
                  <span className="font-black text-xs text-indigo-950 dark:text-indigo-200 uppercase tracking-wider flex items-center gap-1.5">
                    <Smartphone size={16} className="text-indigo-600" />
                    <span>📱 ဤစက်ရှိ ဒေတာ (Local Revision)</span>
                  </span>
                  <span className="text-[9px] bg-indigo-200 dark:bg-indigo-900 text-indigo-900 dark:text-indigo-200 font-mono font-bold px-2 py-0.5 rounded">
                    v{localRecord._version || 1}
                  </span>
                </div>
                <div className="space-y-2 text-xs">
                  {compareFields.map((f, i) => {
                    const localVal = String(localRecord[f.key] || '-');
                    const remoteVal = String(remoteRecord[f.key] || '-');
                    const isDiff = localVal !== remoteVal;
                    return (
                      <div key={i} className={`p-2 rounded-xl transition-all ${isDiff ? 'bg-amber-100 dark:bg-amber-900/40 border border-amber-300 dark:border-amber-700' : 'bg-white/60 dark:bg-slate-800/40'}`}>
                        <div className="text-[10px] font-bold text-slate-400 uppercase">{f.label}</div>
                        <div className={`font-bold ${isDiff ? 'text-amber-950 dark:text-amber-200 font-black' : 'text-slate-800 dark:text-slate-200'}`}>
                          {localVal}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Remote Record Card */}
              <div className="bg-emerald-50/50 dark:bg-emerald-950/30 border-2 border-emerald-200 dark:border-emerald-800 rounded-2xl p-4 space-y-3">
                <div className="flex items-center justify-between border-b border-emerald-200 dark:border-emerald-800 pb-2">
                  <span className="font-black text-xs text-emerald-950 dark:text-emerald-200 uppercase tracking-wider flex items-center gap-1.5">
                    <Server size={16} className="text-emerald-600" />
                    <span>☁️ Cloud / အခြားစက်မှ ဒေတာ (Remote Revision)</span>
                  </span>
                  <span className="text-[9px] bg-emerald-200 dark:bg-emerald-900 text-emerald-900 dark:text-emerald-200 font-mono font-bold px-2 py-0.5 rounded">
                    v{remoteRecord._version || 1}
                  </span>
                </div>
                <div className="space-y-2 text-xs">
                  {compareFields.map((f, i) => {
                    const localVal = String(localRecord[f.key] || '-');
                    const remoteVal = String(remoteRecord[f.key] || '-');
                    const isDiff = localVal !== remoteVal;
                    return (
                      <div key={i} className={`p-2 rounded-xl transition-all ${isDiff ? 'bg-amber-100 dark:bg-amber-900/40 border border-amber-300 dark:border-amber-700' : 'bg-white/60 dark:bg-slate-800/40'}`}>
                        <div className="text-[10px] font-bold text-slate-400 uppercase">{f.label}</div>
                        <div className={`font-bold ${isDiff ? 'text-amber-950 dark:text-amber-200 font-black' : 'text-slate-800 dark:text-slate-200'}`}>
                          {remoteVal}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          </div>

          {/* Action Buttons */}
          <div className="p-6 bg-slate-50 dark:bg-slate-800/80 border-t border-slate-200 dark:border-slate-700/80 flex flex-col sm:flex-row items-center justify-between gap-3 shrink-0">
            <button
              onClick={() => onResolve('KEEP_LOCAL')}
              className="w-full sm:w-auto px-4 py-3 rounded-2xl bg-indigo-600 hover:bg-indigo-700 text-white font-black text-xs uppercase tracking-wider flex items-center justify-center gap-2 shadow-md transition-all cursor-pointer"
            >
              <Smartphone size={16} />
              <span>📱 ဤစက်မှ ဒေတာဖြင့် အတည်ပြုမည် (Keep Local)</span>
            </button>
            <button
              onClick={() => onResolve('KEEP_REMOTE')}
              className="w-full sm:w-auto px-4 py-3 rounded-2xl bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs uppercase tracking-wider flex items-center justify-center gap-2 shadow-md transition-all cursor-pointer"
            >
              <Server size={16} />
              <span>☁️ Cloud ဒေတာဖြင့် အတည်ပြုမည် (Keep Remote)</span>
            </button>
            <button
              onClick={() => onResolve('KEEP_BOTH')}
              className="w-full sm:w-auto px-4 py-3 rounded-2xl bg-slate-900 hover:bg-black text-white font-black text-xs uppercase tracking-wider flex items-center justify-center gap-2 shadow-md transition-all cursor-pointer"
            >
              <Layers size={16} />
              <span>နှစ်ခုလုံး သိမ်းဆည်းမည် (Keep Both Entries)</span>
            </button>
          </div>
        </motion.div>
      </div>
    </AnimatePresence>
  );
};
