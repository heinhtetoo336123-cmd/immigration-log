import React, { useState, useEffect } from 'react';
import { 
  Server, 
  UploadCloud, 
  DownloadCloud, 
  X, 
  CheckCircle2, 
  AlertCircle, 
  Laptop, 
  Smartphone, 
  Tablet, 
  RefreshCw,
  Trash2,
  ShieldCheck,
  ShieldAlert,
  Zap
} from 'lucide-react';
import { motion } from 'framer-motion';
import { 
  CloudStats, 
  fetchCloudCollectionStats, 
  getAllDeletedQueues, 
  purgeDeletedRecordsFromFirestore 
} from '../lib/firebase';
import { ImmRecord } from '../types';

interface CloudSyncStatusModalProps {
  isOpen: boolean;
  onClose: () => void;
  isOnline: boolean;
  isAutoSyncing: boolean;
  isCloudSynced: boolean;
  lastSyncTime: string;
  isQuotaExhausted: boolean;
  cloudAuthUser: {
    username: string;
    role: string;
    deviceName?: string;
    deviceId?: string;
  } | null;
  deviceSessions: any[];
  records: ImmRecord[];
  masterDataCount: number;
  tempRecordsCount: number;
  checkingHistoryCount: number;
  onManualSync: () => Promise<void>;
  onFetchFromCloud: (force?: boolean) => Promise<void>;
  onResetQuota?: () => void;
}

export const CloudSyncStatusModal: React.FC<CloudSyncStatusModalProps> = ({
  isOpen,
  onClose,
  isOnline,
  lastSyncTime,
  isQuotaExhausted,
  cloudAuthUser,
  deviceSessions,
  records,
  onManualSync,
  onFetchFromCloud,
  onResetQuota
}) => {
  const [activeTab, setActiveTab] = useState<'overview' | 'pending' | 'deletions' | 'devices'>('overview');
  const [cloudStats, setCloudStats] = useState<CloudStats | null>(null);
  const [isLoadingStats, setIsLoadingStats] = useState(false);
  const [isSyncingAction, setIsSyncingAction] = useState(false);
  const [statusMessage, setStatusMessage] = useState<string | null>(null);
  const [deletedQueues, setDeletedQueues] = useState<Record<string, string[]>>({});
  const [isPurgingDeletions, setIsPurgingDeletions] = useState(false);

  const pendingRecords = records.filter(r => r.syncStatus === 'pending_sync' || r.syncStatus === 'upload_failed');
  const syncedCount = records.length - pendingRecords.length;

  const totalDeletedPending: number = (Object.values(deletedQueues) as string[][]).reduce((acc: number, curr: string[]) => acc + (curr?.length || 0), 0);

  const loadStats = async () => {
    if (!isOnline) return;
    setIsLoadingStats(true);
    try {
      const stats = await fetchCloudCollectionStats();
      if (stats) setCloudStats(stats);
      setDeletedQueues(getAllDeletedQueues());
    } catch (e) {
      console.warn("Could not load cloud stats:", e);
    } finally {
      setIsLoadingStats(false);
    }
  };

  useEffect(() => {
    if (isOpen) {
      loadStats();
      setDeletedQueues(getAllDeletedQueues());
      setStatusMessage(null);
    }
  }, [isOpen, isOnline]);

  const handleUploadAll = async () => {
    if (isSyncingAction || !isOnline) return;
    setIsSyncingAction(true);
    setStatusMessage("Uploading collections to server...");
    try {
      await onManualSync();
      await loadStats();
      setDeletedQueues(getAllDeletedQueues());
      setStatusMessage("✓ Upload completed successfully.");
    } catch (err: any) {
      setStatusMessage(`Upload failed: ${err?.message || 'Server error'}`);
    } finally {
      setIsSyncingAction(false);
    }
  };

  const handlePullLatest = async () => {
    if (isSyncingAction || !isOnline) return;
    setIsSyncingAction(true);
    setStatusMessage("Fetching latest data from server...");
    try {
      await onFetchFromCloud(true);
      await loadStats();
      setDeletedQueues(getAllDeletedQueues());
      setStatusMessage("✓ Download completed successfully.");
    } catch (err: any) {
      setStatusMessage(`Download failed: ${err?.message || 'Server error'}`);
    } finally {
      setIsSyncingAction(false);
    }
  };

  const handlePurgeAllDeletions = async () => {
    if (isPurgingDeletions || !isOnline) return;
    setIsPurgingDeletions(true);
    setStatusMessage("Purging deleted records from Cloud Server...");
    try {
      let totalPurged = 0;
      const collections = Object.keys(deletedQueues);
      for (const col of collections) {
        const count = await purgeDeletedRecordsFromFirestore(col);
        totalPurged += count;
      }
      await loadStats();
      setDeletedQueues(getAllDeletedQueues());
      setStatusMessage(`✓ Purged ${totalPurged} deleted record(s) from Cloud Server.`);
    } catch (err: any) {
      setStatusMessage(`Purge failed: ${err?.message || 'Server error'}`);
    } finally {
      setIsPurgingDeletions(false);
    }
  };

  const getDeviceIcon = (devName?: string) => {
    const n = (devName || '').toLowerCase();
    if (n.includes('phone') || n.includes('iphone') || n.includes('android') || n.includes('mobile')) {
      return <Smartphone size={14} className="text-slate-400" />;
    }
    if (n.includes('pad') || n.includes('tablet')) {
      return <Tablet size={14} className="text-slate-400" />;
    }
    return <Laptop size={14} className="text-slate-400" />;
  };

  const isUserViewer = cloudAuthUser?.role === 'Viewer';

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs">
      <motion.div
        initial={{ opacity: 0, scale: 0.97 }}
        animate={{ opacity: 1, scale: 1 }}
        exit={{ opacity: 0, scale: 0.97 }}
        className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-xl max-h-[85vh] flex flex-col shadow-xl overflow-hidden text-slate-200"
      >
        {/* Minimalist Header */}
        <div className="px-5 py-4 border-b border-slate-800 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <Server size={18} className="text-slate-400" />
            <h3 className="text-sm font-bold uppercase tracking-wide text-white">
              Cloud Server & Sync Monitor
            </h3>
            <span className={`px-2 py-0.5 rounded text-[10px] font-bold uppercase border ${
              isOnline 
                ? 'bg-emerald-950/60 text-emerald-400 border-emerald-800/60' 
                : 'bg-rose-950/60 text-rose-400 border-rose-800/60'
            }`}>
              {isOnline ? 'Online' : 'Offline'}
            </span>
          </div>
          <button
            onClick={onClose}
            className="text-slate-400 hover:text-white p-1 rounded-lg hover:bg-slate-800 transition-colors cursor-pointer"
          >
            <X size={16} />
          </button>
        </div>

        {/* Minimalist Tab Bar */}
        <div className="flex border-b border-slate-800 bg-slate-900/80 px-4 pt-2 gap-2 text-xs font-semibold overflow-x-auto">
          {[
            { id: 'overview', label: 'Overview' },
            { id: 'pending', label: `Pending (${pendingRecords.length})` },
            { id: 'deletions', label: `Deletions (${totalDeletedPending})` },
            { id: 'devices', label: `Devices (${deviceSessions.length})` }
          ].map(tab => (
            <button
              key={tab.id}
              onClick={() => setActiveTab(tab.id as any)}
              className={`pb-2 px-3 border-b-2 transition-all cursor-pointer whitespace-nowrap ${
                activeTab === tab.id
                  ? 'border-indigo-500 text-white font-bold'
                  : 'border-transparent text-slate-400 hover:text-slate-200'
              }`}
            >
              {tab.label}
            </button>
          ))}
        </div>

        {/* Content Body */}
        <div className="flex-1 overflow-y-auto p-5 space-y-4 text-xs">
          {activeTab === 'overview' && (
            <div className="space-y-4">
              {/* Stat Cards */}
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                <div className="bg-slate-800/40 border border-slate-800 p-3 rounded-xl">
                  <span className="text-[10px] text-slate-400 uppercase font-semibold block">Local Records</span>
                  <div className="text-lg font-bold text-white font-mono mt-0.5">{records.length}</div>
                </div>
                <div className="bg-slate-800/40 border border-slate-800 p-3 rounded-xl">
                  <span className="text-[10px] text-slate-400 uppercase font-semibold block">Synced</span>
                  <div className="text-lg font-bold text-emerald-400 font-mono mt-0.5">{syncedCount}</div>
                </div>
                <div className="bg-slate-800/40 border border-slate-800 p-3 rounded-xl">
                  <span className="text-[10px] text-slate-400 uppercase font-semibold block">Pending Sync</span>
                  <div className={`text-lg font-bold font-mono mt-0.5 ${pendingRecords.length > 0 ? 'text-amber-400' : 'text-slate-400'}`}>
                    {pendingRecords.length}
                  </div>
                </div>
                <div className="bg-slate-800/40 border border-slate-800 p-3 rounded-xl">
                  <span className="text-[10px] text-slate-400 uppercase font-semibold block">Cloud Server</span>
                  <div className="text-lg font-bold text-slate-200 font-mono mt-0.5">
                    {isLoadingStats ? '...' : (cloudStats?.recordsCount ?? '-')}
                  </div>
                </div>
              </div>

              {/* Delete Condition & Sync Effect Panel */}
              <div className="bg-slate-800/30 border border-slate-800 rounded-xl p-3.5 space-y-2.5">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <Trash2 size={14} className="text-rose-400" />
                    <span className="font-bold text-slate-200 uppercase tracking-wide text-[11px]">
                      Delete Condition & Cloud Sync Effect
                    </span>
                  </div>
                  <span className={`px-2 py-0.5 rounded text-[10px] font-mono font-bold ${
                    totalDeletedPending > 0 
                      ? 'bg-amber-950/60 text-amber-300 border border-amber-800/60' 
                      : 'bg-emerald-950/60 text-emerald-400 border border-emerald-800/60'
                  }`}>
                    {totalDeletedPending > 0 ? `${totalDeletedPending} Pending Purge` : 'Cloud Purged'}
                  </span>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 text-[11px] text-slate-300">
                  <div className="bg-slate-900/60 p-2.5 rounded-lg border border-slate-800/80 space-y-1">
                    <div className="flex items-center gap-1.5 text-slate-200 font-semibold">
                      <Zap size={12} className="text-amber-400" /> Direct Cloud Deletion
                    </div>
                    <p className="text-slate-400 text-[10.5px] leading-relaxed">
                      ဖျက်လိုက်သည့် Record အား Firestore မှ တိုက်ရိုက်ထုတ်ပယ်ပြီး Cloud Count အား ချက်ချင်း လျှော့ချပါသည် (ဥပမာ 1748 → 1747)။
                    </p>
                  </div>

                  <div className="bg-slate-900/60 p-2.5 rounded-lg border border-slate-800/80 space-y-1">
                    <div className="flex items-center gap-1.5 text-slate-200 font-semibold">
                      <ShieldCheck size={12} className="text-emerald-400" /> Role & Offline Protection
                    </div>
                    <p className="text-slate-400 text-[10.5px] leading-relaxed">
                      Superadmin/Editor သာ ဖျက်ခွင့်ရှိပြီး Offline ဖျက်မှုများကို Queue ထဲသိမ်းကာ Cloud သို့ ချိတ်ဆက်ချိန်တွင် အလိုအလျောက် Purge ပြုလုပ်ပေးပါသည်။
                    </p>
                  </div>
                </div>

                {isUserViewer && (
                  <div className="p-2 bg-amber-950/30 border border-amber-800/40 rounded-lg text-amber-300 flex items-center gap-2 text-[10.5px]">
                    <ShieldAlert size={13} className="shrink-0 text-amber-400" />
                    <span>လက်ရှိအကောင့်မှာ Viewer ဖြစ်သောကြောင့် လုံခြုံရေးအရ Record ဖျက်ပစ်ခွင့် ပိတ်ထားပါသည်</span>
                  </div>
                )}
              </div>

              {/* Status Message or Quota Alert */}
              {isQuotaExhausted && (
                <div className="p-3 bg-rose-950/30 border border-rose-800/50 rounded-xl text-rose-300 flex items-center justify-between text-[11px]">
                  <div className="flex items-center gap-2">
                    <AlertCircle size={14} className="shrink-0 text-rose-400" />
                    <span>Quota Limit Reached. Data is safely stored in local IndexedDB.</span>
                  </div>
                  {onResetQuota && (
                    <button 
                      onClick={onResetQuota}
                      className="px-2 py-1 bg-rose-900/60 hover:bg-rose-800 text-white rounded text-[10px] font-bold cursor-pointer"
                    >
                      Reset
                    </button>
                  )}
                </div>
              )}

              {statusMessage && (
                <div className="p-2.5 bg-slate-800/60 border border-slate-700/60 rounded-xl text-slate-300 text-[11px]">
                  {statusMessage}
                </div>
              )}

              {/* Action Buttons */}
              <div className="space-y-2 pt-1">
                <div className="flex flex-col sm:flex-row gap-2.5">
                  <button
                    onClick={handleUploadAll}
                    disabled={isSyncingAction || !isOnline}
                    className={`flex-1 py-2.5 px-4 rounded-xl font-bold uppercase text-xs flex items-center justify-center gap-2 transition-all ${
                      !isOnline
                        ? 'bg-slate-800 text-slate-500 cursor-not-allowed border border-slate-800'
                        : isSyncingAction
                        ? 'bg-indigo-700 text-white cursor-wait'
                        : 'bg-indigo-600 hover:bg-indigo-500 text-white cursor-pointer shadow-xs'
                    }`}
                  >
                    <UploadCloud size={15} />
                    <span>{isSyncingAction ? "Processing..." : "UPLOAD ALL TO CLOUD SERVER"}</span>
                  </button>

                  <button
                    onClick={handlePullLatest}
                    disabled={isSyncingAction || !isOnline}
                    className={`flex-1 py-2.5 px-4 rounded-xl font-bold uppercase text-xs flex items-center justify-center gap-2 border transition-all ${
                      !isOnline
                        ? 'bg-slate-900 text-slate-600 border-slate-800 cursor-not-allowed'
                        : 'bg-slate-800 hover:bg-slate-700 text-slate-200 border-slate-700 cursor-pointer'
                    }`}
                  >
                    <DownloadCloud size={15} />
                    <span>PULL LATEST FROM SERVER</span>
                  </button>
                </div>

                <div className="flex items-center justify-between text-[11px] text-slate-500 pt-1">
                  <span>Last Sync: {lastSyncTime || 'None'}</span>
                  <button
                    onClick={loadStats}
                    disabled={isLoadingStats || !isOnline}
                    className="hover:text-slate-300 flex items-center gap-1 cursor-pointer"
                  >
                    <RefreshCw size={11} className={isLoadingStats ? "animate-spin" : ""} /> Check Cloud Info
                  </button>
                </div>
              </div>

              {/* Current Device Summary */}
              <div className="border-t border-slate-800/80 pt-3 flex items-center justify-between text-[11px] text-slate-400">
                <div className="flex items-center gap-2">
                  {getDeviceIcon(cloudAuthUser?.deviceName)}
                  <span>{cloudAuthUser?.deviceName || 'Local Duty Terminal'}</span>
                  <span className="text-[10px] bg-slate-800 px-1.5 py-0.2 rounded font-mono text-slate-300">
                    {cloudAuthUser?.role || 'Editor'}
                  </span>
                </div>
                <span className="font-mono text-slate-500">{cloudAuthUser?.username || 'OFFICER'}</span>
              </div>
            </div>
          )}

          {activeTab === 'pending' && (
            <div className="space-y-3">
              {pendingRecords.length === 0 ? (
                <div className="py-8 text-center text-slate-500 space-y-1.5">
                  <CheckCircle2 size={24} className="mx-auto text-emerald-500/60" />
                  <div className="text-xs font-semibold text-slate-400">No pending uploads</div>
                  <div className="text-[11px]">All records are in sync with cloud storage.</div>
                </div>
              ) : (
                <div className="space-y-1.5 max-h-72 overflow-y-auto pr-1">
                  {pendingRecords.map((r, i) => (
                    <div key={r.id || i} className="p-2 bg-slate-800/40 border border-slate-800 rounded-lg flex items-center justify-between text-[11px]">
                      <div className="flex items-center gap-2">
                        <span className="font-mono text-slate-400">#{r.id}</span>
                        <span className="font-semibold text-slate-200">{r.fullname || r.passport || 'Record'}</span>
                      </div>
                      <span className="text-[10px] px-2 py-0.5 rounded bg-amber-950/60 text-amber-400 border border-amber-800/60">
                        {r.syncStatus || 'pending_sync'}
                      </span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {activeTab === 'deletions' && (
            <div className="space-y-3">
              <div className="p-3 bg-slate-800/30 border border-slate-800 rounded-xl flex items-center justify-between">
                <div>
                  <span className="text-xs font-bold text-white block">Pending Deletion Queue</span>
                  <span className="text-[11px] text-slate-400">
                    {totalDeletedPending > 0 
                      ? `${totalDeletedPending} ID(s) queued for remote purge`
                      : 'No deleted records waiting to purge'}
                  </span>
                </div>
                {totalDeletedPending > 0 && (
                  <button
                    onClick={handlePurgeAllDeletions}
                    disabled={isPurgingDeletions || !isOnline}
                    className="px-3 py-1.5 bg-rose-600 hover:bg-rose-500 text-white rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 cursor-pointer"
                  >
                    <Trash2 size={13} />
                    <span>{isPurgingDeletions ? 'Purging...' : 'Purge from Cloud'}</span>
                  </button>
                )}
              </div>

              {totalDeletedPending === 0 ? (
                <div className="py-8 text-center text-slate-500 space-y-1.5">
                  <CheckCircle2 size={24} className="mx-auto text-emerald-500/60" />
                  <div className="text-xs font-semibold text-slate-400">Deletion Queue is Clear</div>
                  <div className="text-[11px]">All local deletions have been purged from Firestore.</div>
                </div>
              ) : (
                <div className="space-y-2 max-h-64 overflow-y-auto pr-1">
                  {(Object.entries(deletedQueues) as [string, string[]][]).map(([col, ids]) => (
                    Array.isArray(ids) && ids.length > 0 && (
                      <div key={col} className="bg-slate-800/40 border border-slate-800 rounded-lg p-2.5 space-y-1.5">
                        <div className="flex items-center justify-between text-[11px] font-bold text-slate-300">
                          <span className="uppercase">{col}</span>
                          <span className="text-rose-400">{ids.length} deleted</span>
                        </div>
                        <div className="flex flex-wrap gap-1">
                          {ids.map(id => (
                            <span key={id} className="text-[10px] font-mono px-2 py-0.5 bg-slate-900 text-rose-300 border border-slate-700 rounded">
                              ID: {id}
                            </span>
                          ))}
                        </div>
                      </div>
                    )
                  ))}
                </div>
              )}
            </div>
          )}

          {activeTab === 'devices' && (
            <div className="space-y-2 max-h-72 overflow-y-auto pr-1">
              {deviceSessions.length === 0 ? (
                <div className="py-8 text-center text-slate-500">
                  <div className="text-xs">No connected device records found</div>
                </div>
              ) : (
                deviceSessions.map(session => (
                  <div key={session.deviceId} className="p-2.5 bg-slate-800/40 border border-slate-800 rounded-xl flex items-center justify-between text-[11px]">
                    <div className="flex items-center gap-2.5">
                      {getDeviceIcon(session.deviceName)}
                      <div>
                        <div className="font-semibold text-slate-200">{session.deviceName || 'Terminal'}</div>
                        <div className="text-[10px] text-slate-500">{session.username || 'User'} • {session.lastActive || 'Active'}</div>
                      </div>
                    </div>
                    <span className="text-[10px] bg-slate-800 px-2 py-0.5 rounded text-slate-300 font-mono">
                      {session.accountRole || 'Editor'}
                    </span>
                  </div>
                ))
              )}
            </div>
          )}
        </div>
      </motion.div>
    </div>
  );
};
