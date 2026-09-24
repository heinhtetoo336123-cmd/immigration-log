import React, { useState, useEffect } from 'react';
import { 
  Cloud, 
  CloudCheck, 
  RefreshCw, 
  Smartphone, 
  Laptop, 
  Tablet, 
  CheckCircle2, 
  AlertCircle, 
  Clock, 
  Database, 
  UploadCloud, 
  DownloadCloud, 
  X, 
  Wifi, 
  WifiOff, 
  ShieldCheck, 
  ArrowUpRight,
  Server
} from 'lucide-react';
import { motion, AnimatePresence } from 'framer-motion';
import { CloudStats, fetchCloudCollectionStats } from '../lib/firebase';
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
  isAutoSyncing,
  isCloudSynced,
  lastSyncTime,
  isQuotaExhausted,
  cloudAuthUser,
  deviceSessions,
  records,
  masterDataCount,
  tempRecordsCount,
  checkingHistoryCount,
  onManualSync,
  onFetchFromCloud,
  onResetQuota
}) => {
  const [activeTab, setActiveTab] = useState<'status' | 'pending' | 'devices'>('status');
  const [cloudStats, setCloudStats] = useState<CloudStats | null>(null);
  const [isLoadingStats, setIsLoadingStats] = useState(false);
  const [isSyncingAction, setIsSyncingAction] = useState(false);

  // Compute pending vs synced records
  const pendingRecords = records.filter(r => r.syncStatus === 'pending_sync' || r.syncStatus === 'upload_failed');
  const syncedRecordsCount = records.length - pendingRecords.length;

  const loadStats = async () => {
    if (!isOnline) return;
    setIsLoadingStats(true);
    try {
      const stats = await fetchCloudCollectionStats();
      if (stats) setCloudStats(stats);
    } catch (e) {
      console.warn("Could not load cloud stats:", e);
    } finally {
      setIsLoadingStats(false);
    }
  };

  useEffect(() => {
    if (isOpen) {
      loadStats();
    }
  }, [isOpen, isOnline]);

  const handleTriggerUpload = async () => {
    if (isSyncingAction) return;
    setIsSyncingAction(true);
    try {
      await onManualSync();
      await loadStats();
    } finally {
      setIsSyncingAction(false);
    }
  };

  const handleTriggerDownload = async () => {
    if (isSyncingAction) return;
    setIsSyncingAction(true);
    try {
      await onFetchFromCloud(true);
      await loadStats();
    } finally {
      setIsSyncingAction(false);
    }
  };

  const getDeviceIcon = (devName?: string) => {
    const n = (devName || '').toLowerCase();
    if (n.includes('phone') || n.includes('iphone') || n.includes('android') || n.includes('mobile')) {
      return <Smartphone size={16} className="text-emerald-400" />;
    }
    if (n.includes('pad') || n.includes('tablet')) {
      return <Tablet size={16} className="text-cyan-400" />;
    }
    return <Laptop size={16} className="text-indigo-400" />;
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-slate-950/75 backdrop-blur-sm animate-fade-in">
      <motion.div
        initial={{ opacity: 0, scale: 0.95, y: 10 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.95, y: 10 }}
        className="bg-slate-900 border border-slate-700/80 rounded-3xl w-full max-w-2xl max-h-[90vh] flex flex-col shadow-2xl overflow-hidden text-slate-100"
      >
        {/* Header */}
        <div className="p-4 sm:p-5 bg-gradient-to-r from-slate-900 via-indigo-950/40 to-slate-900 border-b border-slate-800 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-2xl bg-indigo-600/30 border border-indigo-500/40 flex items-center justify-center text-indigo-400 shadow-inner">
              <Server size={20} />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h3 className="text-sm sm:text-base font-black uppercase tracking-tight text-white">
                  Cloud Server & Sync Monitor
                </h3>
                <span className={`px-2 py-0.5 rounded-full text-[9px] font-black uppercase tracking-wider border ${
                  isOnline 
                    ? 'bg-emerald-500/20 text-emerald-300 border-emerald-500/30' 
                    : 'bg-rose-500/20 text-rose-300 border-rose-500/30'
                }`}>
                  {isOnline ? 'Online' : 'Offline'}
                </span>
              </div>
              <p className="text-[11px] text-slate-400 font-medium">
                ဆာဗာဒေတာနှင့် အခြားစက်များ (ဖုန်း/ကွန်ပျူတာ) အချက်အလက် စစ်ဆေးမှု စနစ်
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="w-8 h-8 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white flex items-center justify-center transition-colors cursor-pointer"
          >
            <X size={16} />
          </button>
        </div>

        {/* Tab Selection */}
        <div className="flex border-b border-slate-800 bg-slate-900/60 p-1.5 gap-1.5 text-xs font-bold">
          <button
            onClick={() => setActiveTab('status')}
            className={`flex-1 py-2 px-3 rounded-xl transition-all flex items-center justify-center gap-1.5 cursor-pointer ${
              activeTab === 'status'
                ? 'bg-indigo-600 text-white shadow-md'
                : 'text-slate-400 hover:text-white hover:bg-slate-800/60'
            }`}
          >
            <Cloud size={14} />
            <span>ဆာဗာ ချိတ်ဆက်မှု (Sync Status)</span>
          </button>
          <button
            onClick={() => setActiveTab('pending')}
            className={`flex-1 py-2 px-3 rounded-xl transition-all flex items-center justify-center gap-1.5 cursor-pointer ${
              activeTab === 'pending'
                ? 'bg-indigo-600 text-white shadow-md'
                : 'text-slate-400 hover:text-white hover:bg-slate-800/60'
            }`}
          >
            <Clock size={14} />
            <span>မရောက်သေးသော ဒေတာ ({pendingRecords.length})</span>
          </button>
          <button
            onClick={() => setActiveTab('devices')}
            className={`flex-1 py-2 px-3 rounded-xl transition-all flex items-center justify-center gap-1.5 cursor-pointer ${
              activeTab === 'devices'
                ? 'bg-indigo-600 text-white shadow-md'
                : 'text-slate-400 hover:text-white hover:bg-slate-800/60'
            }`}
          >
            <Smartphone size={14} />
            <span>ချိတ်ဆက်ထားသော စက်များ ({deviceSessions.length})</span>
          </button>
        </div>

        {/* Body Content */}
        <div className="flex-1 overflow-y-auto p-4 sm:p-5 space-y-4 text-xs">
          {activeTab === 'status' && (
            <div className="space-y-4">
              {/* Highlight Status Card */}
              <div className={`p-4 rounded-2xl border ${
                pendingRecords.length === 0
                  ? 'bg-emerald-950/20 border-emerald-500/30 text-emerald-100'
                  : 'bg-amber-950/20 border-amber-500/30 text-amber-100'
              }`}>
                <div className="flex items-start gap-3">
                  <div className={`p-2 rounded-xl mt-0.5 ${
                    pendingRecords.length === 0 ? 'bg-emerald-500/20 text-emerald-400' : 'bg-amber-500/20 text-amber-400'
                  }`}>
                    {pendingRecords.length === 0 ? <CheckCircle2 size={20} /> : <AlertCircle size={20} />}
                  </div>
                  <div className="flex-1">
                    <h4 className="font-black text-sm uppercase">
                      {pendingRecords.length === 0
                        ? '✓ စာရင်းများအားလုံး ဆာဗာပေါ်သို့ ရောက်ရှိသိမ်းဆည်းပြီးပါပြီ'
                        : `⚠️ ဆာဗာသို့ မရောက်သေးသော ဒေတာ (${pendingRecords.length}) ခု ရှိနေပါသည်`}
                    </h4>
                    <p className="text-[11px] text-slate-300 mt-1">
                      {pendingRecords.length === 0
                        ? `လက်ရှိစက်ရှိ မှတ်တမ်း (${records.length}) ခုစလုံးသည် Google Cloud Server ပေါ်တွင် လုံခြုံစွာ သိမ်းဆည်းပြီးဖြစ်၍ အခြားစက်များမှလည်း တိုက်ရိုက် ကြည့်ရှုနိုင်ပါသည်။`
                        : 'အင်တာနက်လိုင်း ချို့ယွင်းမှု သို့မဟုတ် မပို့ရသေးသော စာရင်းများ ရှိနေပါသည်။ အောက်ပါ "Upload All Pending" ခလုတ်ကို နှိပ်၍ ဆာဗာသို့ ချက်ချင်း ပို့ဆောင်နိုင်ပါသည်။'}
                    </p>
                    <div className="flex items-center gap-2 mt-2 text-[10px] text-slate-400 font-mono">
                      <span>နောက်ဆုံး Sync အချိန်:</span>
                      <span className="font-bold text-indigo-300">{lastSyncTime || 'မပြုလုပ်ရသေးပါ'}</span>
                    </div>
                  </div>
                </div>
              </div>

              {/* Data Counters Grid */}
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5">
                <div className="bg-slate-800/70 border border-slate-700/60 p-3 rounded-2xl">
                  <span className="text-[10px] font-bold text-slate-400 uppercase block">Total Local Recs</span>
                  <div className="text-xl font-black text-white mt-1 font-mono">{records.length}</div>
                  <span className="text-[9px] text-indigo-300 font-semibold">ဤစက်ရှိ စာရင်းစုစုပေါင်း</span>
                </div>
                <div className="bg-slate-800/70 border border-slate-700/60 p-3 rounded-2xl">
                  <span className="text-[10px] font-bold text-emerald-400 uppercase block">Server Synced</span>
                  <div className="text-xl font-black text-emerald-400 mt-1 font-mono">{syncedRecordsCount}</div>
                  <span className="text-[9px] text-emerald-300/80 font-semibold">ဆာဗာပေါ် ရောက်ပြီး</span>
                </div>
                <div className="bg-slate-800/70 border border-slate-700/60 p-3 rounded-2xl">
                  <span className="text-[10px] font-bold text-amber-400 uppercase block">Pending Upload</span>
                  <div className="text-xl font-black text-amber-400 mt-1 font-mono">{pendingRecords.length}</div>
                  <span className="text-[9px] text-amber-300/80 font-semibold">ဆာဗာ မရောက်သေး</span>
                </div>
                <div className="bg-slate-800/70 border border-slate-700/60 p-3 rounded-2xl">
                  <span className="text-[10px] font-bold text-cyan-400 uppercase block">Cloud Server Recs</span>
                  <div className="text-xl font-black text-cyan-400 mt-1 font-mono">
                    {isLoadingStats ? '...' : (cloudStats?.recordsCount ?? '-')}
                  </div>
                  <span className="text-[9px] text-cyan-300/80 font-semibold">ဆာဗာရှိ စုစုပေါင်း</span>
                </div>
              </div>

              {/* Unified Sync Actions */}
              <div className="bg-slate-800/50 border border-slate-700/70 rounded-2xl p-4 space-y-3">
                <div className="flex items-center justify-between">
                  <span className="font-black text-xs uppercase text-slate-200 flex items-center gap-1.5">
                    <Database size={14} className="text-indigo-400" />
                    Unified Cloud Synchronization (ဒေတာအားလုံး တပြိုင်နက် ချိတ်ဆက်ခြင်း)
                  </span>
                  <button
                    onClick={loadStats}
                    disabled={isLoadingStats}
                    className="text-[10px] text-indigo-400 hover:text-indigo-300 flex items-center gap-1 cursor-pointer"
                  >
                    <RefreshCw size={10} className={isLoadingStats ? "animate-spin" : ""} /> Check Cloud
                  </button>
                </div>
                <p className="text-[11px] text-slate-400">
                  ဤနေရာမှတစ်ဆင့် Flight/Foreigner Entries (FFE)၊ Master Data၊ Checkpoint နှင့် History စာရင်းအားလုံးကို Google Cloud Server နှင့် တပြိုင်နက် Upload/Download Sync ပြုလုပ်နိုင်ပါသည်။
                </p>

                <div className="flex flex-col sm:flex-row gap-2 pt-1">
                  <button
                    onClick={handleTriggerUpload}
                    disabled={isSyncingAction || !isOnline}
                    className={`flex-1 py-2.5 px-4 rounded-xl font-black uppercase text-xs flex items-center justify-center gap-2 shadow-lg transition-all ${
                      !isOnline
                        ? 'bg-slate-700 text-slate-400 cursor-not-allowed'
                        : isSyncingAction
                        ? 'bg-indigo-700 text-white cursor-wait'
                        : 'bg-indigo-600 hover:bg-indigo-500 text-white cursor-pointer shadow-indigo-900/30'
                    }`}
                  >
                    <UploadCloud size={16} className={isSyncingAction ? "animate-bounce" : ""} />
                    <span>{isSyncingAction ? "Uploading to Server..." : "Upload All to Cloud Server (ဆာဗာသို့ ပို့မည်)"}</span>
                  </button>

                  <button
                    onClick={handleTriggerDownload}
                    disabled={isSyncingAction || !isOnline}
                    className={`flex-1 py-2.5 px-4 rounded-xl font-black uppercase text-xs flex items-center justify-center gap-2 border transition-all ${
                      !isOnline
                        ? 'bg-slate-800 text-slate-500 border-slate-700 cursor-not-allowed'
                        : 'bg-slate-800 hover:bg-slate-700 text-cyan-300 border-slate-700 hover:border-cyan-500/40 cursor-pointer'
                    }`}
                  >
                    <DownloadCloud size={16} />
                    <span>Pull Latest from Server (ဆာဗာမှ ဒေတာ ရယူမည်)</span>
                  </button>
                </div>
              </div>

              {/* Current Device Info */}
              <div className="bg-slate-800/40 border border-slate-700/50 p-3 rounded-2xl flex items-center justify-between text-[11px]">
                <div className="flex items-center gap-2.5">
                  <div className="p-2 rounded-xl bg-indigo-500/20 text-indigo-300">
                    {getDeviceIcon(cloudAuthUser?.deviceName)}
                  </div>
                  <div>
                    <div className="font-bold text-white flex items-center gap-1.5">
                      <span>{cloudAuthUser?.deviceName || 'This Device'}</span>
                      <span className="text-[9px] bg-indigo-500/20 text-indigo-300 px-1.5 py-0.2 rounded border border-indigo-400/30 font-bold">
                        {cloudAuthUser?.role || 'Editor'}
                      </span>
                    </div>
                    <span className="text-[10px] text-slate-400 font-mono">
                      User: {cloudAuthUser?.username || 'OFFICER'} • ID: {cloudAuthUser?.deviceId?.substring(0, 14) || 'local'}...
                    </span>
                  </div>
                </div>
                <div className="text-right">
                  <span className="text-[10px] text-emerald-400 font-bold block">● Connected</span>
                  <span className="text-[9px] text-slate-400">Active Duty</span>
                </div>
              </div>
            </div>
          )}

          {activeTab === 'pending' && (
            <div className="space-y-3">
              <div className="flex items-center justify-between">
                <div>
                  <h4 className="font-black text-xs uppercase text-slate-200">
                    ဆာဗာသို့ မရောက်ရှိသေးသော ဒေတာများ (Pending Upload Records)
                  </h4>
                  <p className="text-[11px] text-slate-400">
                    ဤစာရင်းများသည် ဖုန်း/စက်ထဲတွင်သာ ရှိနေပြီး Server သို့ မရောက်သေးသော စာရင်းများဖြစ်ပါသည်
                  </p>
                </div>
                {pendingRecords.length > 0 && (
                  <button
                    onClick={handleTriggerUpload}
                    disabled={isSyncingAction || !isOnline}
                    className="bg-indigo-600 hover:bg-indigo-500 text-white font-black text-[11px] px-3 py-1.5 rounded-xl uppercase flex items-center gap-1 cursor-pointer shadow-md"
                  >
                    <UploadCloud size={13} /> Upload All Now
                  </button>
                )}
              </div>

              {pendingRecords.length === 0 ? (
                <div className="p-8 text-center bg-slate-800/30 border border-slate-700/50 rounded-2xl space-y-2">
                  <CheckCircle2 size={32} className="mx-auto text-emerald-400" />
                  <div className="font-black text-sm text-slate-200">မရောက်သေးသော စာရင်း လုံးဝမရှိပါ</div>
                  <p className="text-slate-400 text-xs">
                    မှတ်တမ်းအားလုံးသည် Cloud Server ပေါ်သို့ ၁၀၀% အောင်မြင်စွာ ရောက်ရှိပြီး ဖြစ်ပါသည်။
                  </p>
                </div>
              ) : (
                <div className="space-y-2 max-h-80 overflow-y-auto pr-1">
                  {pendingRecords.map((r, idx) => (
                    <div
                      key={r.id || idx}
                      className="bg-slate-800/80 border border-slate-700/80 p-3 rounded-2xl flex items-center justify-between gap-3"
                    >
                      <div className="min-w-0">
                        <div className="flex items-center gap-2">
                          <span className="font-mono font-black text-xs text-indigo-300">{r.passport}</span>
                          <span className={`text-[9px] px-1.5 py-0.2 rounded font-bold uppercase ${
                            r.mode === 'IN' ? 'bg-emerald-500/20 text-emerald-300' : 'bg-orange-500/20 text-orange-300'
                          }`}>
                            {r.mode}
                          </span>
                          <span className="text-[9px] bg-amber-500/20 text-amber-300 border border-amber-500/30 px-1.5 py-0.2 rounded font-bold flex items-center gap-1">
                            <Clock size={8} /> Local Only
                          </span>
                        </div>
                        <div className="text-xs font-bold text-white truncate mt-0.5">{r.fullname}</div>
                        <div className="text-[10px] text-slate-400 font-mono">
                          {r.timestamp} • {r.nationality || '-'} • Flight/Vehicle: {r.vehicleInfo || '-'}
                        </div>
                      </div>
                      <button
                        onClick={handleTriggerUpload}
                        disabled={isSyncingAction}
                        className="px-2.5 py-1 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-[10px] font-black uppercase shrink-0 flex items-center gap-1 cursor-pointer"
                      >
                        <UploadCloud size={10} /> Sync
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {activeTab === 'devices' && (
            <div className="space-y-3">
              <div>
                <h4 className="font-black text-xs uppercase text-slate-200">
                  ချိတ်ဆက်အသုံးပြုနေသော စက်များ (Connected Devices)
                </h4>
                <p className="text-[11px] text-slate-400">
                  ဆာဗာသို့ ချိတ်ဆက်ထားသော အခြားစက်များ (ဖုန်း၊ တက်ဘလက်၊ ကွန်ပျူတာ) ၏ နောက်ဆုံး လှုပ်ရှားမှု အခြေအနေ
                </p>
              </div>

              <div className="space-y-2">
                {deviceSessions.map((session, idx) => {
                  const isCurrent = session.deviceId === cloudAuthUser?.deviceId;
                  const isRecentlyActive = session.lastPingTimestamp && (Date.now() - session.lastPingTimestamp < 2 * 60 * 60 * 1000);

                  return (
                    <div
                      key={session.deviceId || idx}
                      className={`p-3 rounded-2xl border flex items-center justify-between gap-3 ${
                        isCurrent 
                          ? 'bg-indigo-950/30 border-indigo-500/40' 
                          : 'bg-slate-800/60 border-slate-700/60'
                      }`}
                    >
                      <div className="flex items-center gap-3 min-w-0">
                        <div className="p-2 rounded-xl bg-slate-700/60 shrink-0">
                          {getDeviceIcon(session.deviceName)}
                        </div>
                        <div className="min-w-0">
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <span className="font-black text-white text-xs truncate">
                              {session.deviceName || 'Device'}
                            </span>
                            {isCurrent && (
                              <span className="bg-indigo-500 text-white text-[8px] px-1.5 py-0.2 rounded font-black uppercase">
                                This Device
                              </span>
                            )}
                            <span className="text-[9px] bg-slate-700 text-slate-300 px-1.5 py-0.2 rounded font-bold">
                              {session.accountRole || 'Editor'}
                            </span>
                          </div>
                          <div className="text-[10px] text-slate-400 font-mono mt-0.5 truncate">
                            User: {session.username || 'OFFICER'} • Last Active: {session.lastActive || 'Just now'}
                          </div>
                        </div>
                      </div>

                      <div className="text-right shrink-0">
                        {isRecentlyActive ? (
                          <span className="text-[10px] text-emerald-400 font-bold flex items-center gap-1 justify-end">
                            <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse" />
                            Online
                          </span>
                        ) : (
                          <span className="text-[10px] text-slate-400 font-bold flex items-center gap-1 justify-end">
                            <span className="w-1.5 h-1.5 rounded-full bg-slate-500" />
                            Idle
                          </span>
                        )}
                        <span className="text-[9px] text-slate-500 font-mono block">
                          ID: {session.deviceId ? session.deviceId.substring(0, 10) : 'dev'}...
                        </span>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="p-3.5 sm:p-4 bg-slate-900 border-t border-slate-800 flex items-center justify-between text-xs">
          <div className="flex items-center gap-2 text-slate-400 text-[11px]">
            <ShieldCheck size={14} className="text-indigo-400" />
            <span>Google Firebase Enterprise Encrypted Sync</span>
          </div>
          <button
            onClick={onClose}
            className="px-4 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white rounded-xl font-black uppercase text-[11px] transition-colors cursor-pointer"
          >
            Close
          </button>
        </div>
      </motion.div>
    </div>
  );
};
