import React, { useState, useEffect, useMemo } from 'react';
import { Trash2, RefreshCw, Search, RotateCcw, ShieldAlert, FileText, Filter, Calendar, User, Smartphone, AlertCircle, CheckCircle2 } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { DeletedRecord } from '../types';
import { fetchDeletedRecords, restoreDeletedRecord } from '../lib/firebase';

interface RecycleBinViewProps {
  currentUser: any;
  showToast: (msg: string) => void;
  onRestored: (restoredInfo: { item: any; collection: string }) => void;
  onRequestDeleteConfirm: (title: string, itemDetails: any, onConfirm: () => void) => void;
}

export const RecycleBinView: React.FC<RecycleBinViewProps> = ({
  currentUser,
  showToast,
  onRestored,
  onRequestDeleteConfirm
}) => {
  const [deletedRecords, setDeletedRecords] = useState<DeletedRecord[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const [filterCollection, setFilterCollection] = useState<string>('ALL');
  const [isLoading, setIsLoading] = useState(false);

  const loadDeleted = async () => {
    setIsLoading(true);
    try {
      const records = await fetchDeletedRecords();
      setDeletedRecords(records || []);
    } catch (e) {
      console.error(e);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    loadDeleted();
  }, []);

  const handleRestore = async (rec: DeletedRecord) => {
    const activeOfficer = currentUser ? `${currentUser.title} ${currentUser.name}` : 'Duty Officer';
    showToast(`🔄 "${rec.title}" အား မူရင်းနေရာသို့ ပြန်လည်ရယူနေပါသည်...`);
    const result = await restoreDeletedRecord(rec.id, activeOfficer);
    if (result) {
      setDeletedRecords(prev => prev.filter(d => d.id !== rec.id));
      onRestored(result);
      showToast(`✓ "${rec.title}" အား အောင်မြင်စွာ မူရင်းစနစ်သို့ ပြန်လည်ရယူပြီးပါပြီ (Restored ✓)`);
    } else {
      showToast('❌ ပြန်လည်ရယူရာတွင် အမှားအယွင်းရှိပါသည်');
    }
  };

  const filteredList = useMemo(() => {
    return deletedRecords.filter(item => {
      if (filterCollection !== 'ALL' && item.originalCollection !== filterCollection) {
        return false;
      }
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const matchTitle = (item.title || '').toLowerCase().includes(q);
        const matchSubtitle = (item.subtitle || '').toLowerCase().includes(q);
        const matchOfficer = (item.deletedBy || '').toLowerCase().includes(q);
        const matchReason = (item.reason || '').toLowerCase().includes(q);
        return matchTitle || matchSubtitle || matchOfficer || matchReason;
      }
      return true;
    });
  }, [deletedRecords, filterCollection, searchQuery]);

  return (
    <div className="max-w-7xl mx-auto py-8 px-4 space-y-6 select-text no-print">
      {/* Header Banner */}
      <div className="bg-rose-950 text-white p-6 sm:p-8 rounded-3xl shadow-xl border border-rose-800/80 flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
        <div className="flex items-center gap-4">
          <div className="w-14 h-14 rounded-2xl bg-rose-900/80 text-rose-200 flex items-center justify-center shrink-0 border border-rose-700/50 shadow-inner">
            <Trash2 size={32} />
          </div>
          <div>
            <div className="flex items-center gap-2 flex-wrap">
              <h2 className="text-xl sm:text-2xl font-black uppercase tracking-tight text-white">
                Recycle Bin / Soft-Delete Ledger (အမှိုက်ပုံး)
              </h2>
              <span className="bg-rose-900 text-rose-200 px-2.5 py-0.5 rounded-full text-xs font-mono font-bold border border-rose-700">
                {deletedRecords.length} Archived
              </span>
            </div>
            <p className="text-xs text-rose-300 font-medium mt-1">
              ပယ်ဖျက်လိုက်သည့် မှတ်တမ်းများအား ၁၀၀% မပျောက်ပျက်စေရန် အလိုအလျောက် ထိန်းသိမ်းထားပြီး မူရင်းနေရာသို့ အချိန်မရွေး ပြန်လည်ရယူနိုင်ပါသည်။
            </p>
          </div>
        </div>

        <button
          onClick={loadDeleted}
          disabled={isLoading}
          className="px-4 py-2.5 bg-rose-900/60 hover:bg-rose-900 text-white font-black text-xs uppercase tracking-wider rounded-2xl border border-rose-700 flex items-center gap-2 transition-all cursor-pointer shadow-sm"
        >
          <RefreshCw size={14} className={isLoading ? 'animate-spin' : ''} />
          <span>Refresh Bin</span>
        </button>
      </div>

      {/* Control Toolbar */}
      <div className="bg-white dark:bg-slate-900 p-4 rounded-2xl shadow-md border border-slate-200 dark:border-slate-800 flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-4">
        {/* Search Field */}
        <div className="relative flex-1">
          <Search size={16} className="absolute left-3.5 top-3.5 text-slate-400" />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Search by passport, name, officer name, or reason..."
            className="w-full pl-10 pr-4 py-2.5 rounded-xl bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 text-xs font-semibold focus:outline-none focus:ring-2 focus:ring-rose-500"
          />
        </div>

        {/* Collection Filter Bar */}
        <div className="flex items-center gap-1.5 overflow-x-auto pb-1 sm:pb-0">
          {[
            { id: 'ALL', label: 'All Items' },
            { id: 'records', label: 'Flight Logs (FFE)' },
            { id: 'masterData', label: 'Master DB' },
            { id: 'watchList', label: 'WatchList' },
            { id: 'dossierHistory', label: 'Dossiers' },
          ].map((cat) => (
            <button
              key={cat.id}
              onClick={() => setFilterCollection(cat.id)}
              className={`px-3 py-1.5 text-xs font-extrabold rounded-xl transition-all cursor-pointer whitespace-nowrap ${
                filterCollection === cat.id
                  ? 'bg-rose-600 text-white shadow-sm'
                  : 'bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-400 hover:bg-slate-200'
              }`}
            >
              {cat.label}
            </button>
          ))}
        </div>
      </div>

      {/* Deleted Items List */}
      <div className="space-y-3">
        {filteredList.length === 0 ? (
          <div className="p-12 text-center bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl space-y-3">
            <Trash2 size={40} className="mx-auto text-slate-300 dark:text-slate-700" />
            <div className="text-sm font-black text-slate-500 dark:text-slate-400 uppercase">
              No soft-deleted records found in Recycle Bin
            </div>
            <p className="text-xs text-slate-400 max-w-sm mx-auto">
              ပယ်ဖျက်ထားသော မှတ်တမ်းများ မရှိသေးပါ သို့မဟုတ် ရွေးချယ်ထားသော Filter ဖြင့် ကိုက်ညီသည့် ဒေတာ မတွေ့ရှိပါ။
            </p>
          </div>
        ) : (
          filteredList.map((item) => (
            <motion.div
              key={item.id}
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 hover:border-rose-300 dark:hover:border-rose-800 p-4 sm:p-5 rounded-2xl shadow-sm transition-all flex flex-col md:flex-row items-start md:items-center justify-between gap-4"
            >
              <div className="space-y-1.5 flex-1 min-w-0">
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="px-2 py-0.5 rounded text-[9px] font-black uppercase bg-rose-100 text-rose-800 dark:bg-rose-950 dark:text-rose-200 border border-rose-300 dark:border-rose-800">
                    {item.originalCollection}
                  </span>
                  <h4 className="text-sm font-black text-slate-900 dark:text-white uppercase truncate">
                    {item.title}
                  </h4>
                </div>

                {item.subtitle && (
                  <div className="text-xs font-semibold text-slate-600 dark:text-slate-400 font-mono">
                    {item.subtitle}
                  </div>
                )}

                <div className="flex flex-wrap items-center gap-x-4 gap-y-1 text-[11px] font-bold text-slate-500 dark:text-slate-400 pt-1">
                  <span className="flex items-center gap-1">
                    <Calendar size={12} className="text-rose-500" />
                    <span>Deleted At: {new Date(item.deletedAt).toLocaleString()}</span>
                  </span>
                  <span className="flex items-center gap-1">
                    <User size={12} className="text-indigo-500" />
                    <span>Deleted By: {item.deletedBy} {item.deletedByRole ? `(${item.deletedByRole})` : ''}</span>
                  </span>
                  {item.reason && (
                    <span className="bg-slate-100 dark:bg-slate-800 px-2 py-0.5 rounded text-[10px] text-slate-700 dark:text-slate-300 italic">
                      Reason: {item.reason}
                    </span>
                  )}
                  {item.deviceId && (
                    <span className="flex items-center gap-1 text-[10px] font-mono">
                      <Smartphone size={11} /> {item.deviceId}
                    </span>
                  )}
                </div>
              </div>

              {/* Restore Action */}
              <div className="shrink-0 w-full md:w-auto flex items-center justify-end">
                <button
                  onClick={() => handleRestore(item)}
                  className="w-full md:w-auto px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs uppercase tracking-wider shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer active:scale-95"
                >
                  <RotateCcw size={15} />
                  <span>Restore Record (ပြန်လည်ရယူမည်)</span>
                </button>
              </div>
            </motion.div>
          ))
        )}
      </div>
    </div>
  );
};
