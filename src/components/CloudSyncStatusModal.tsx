import React, { useState, useEffect, useMemo } from 'react';
import { 
  Cloud, 
  RefreshCw, 
  Smartphone, 
  Laptop, 
  Tablet, 
  CheckCircle2, 
  Clock, 
  Database, 
  UploadCloud, 
  DownloadCloud, 
  X, 
  Server,
  Layers,
  FileSpreadsheet,
  ShieldAlert,
  Plane,
  Truck,
  Check
} from 'lucide-react';
import { motion } from 'framer-motion';
import { CloudStats, fetchCloudCollectionStats } from '../lib/firebase';
import { ImmRecord, MasterItem, VehicleSummary, DossierRecord, WatchListRecord } from '../types';

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
  tempRecords?: ImmRecord[];
  masterData?: MasterItem[];
  vehicleSummaries?: VehicleSummary[];
  dossierHistory?: DossierRecord[];
  watchList?: WatchListRecord[];
  onManualSync: () => Promise<void>;
  onFetchFromCloud: (force?: boolean, silent?: boolean) => Promise<void>;
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
  deviceSessions = [],
  records = [],
  tempRecords = [],
  masterData = [],
  vehicleSummaries = [],
  dossierHistory = [],
  watchList = [],
  onManualSync,
  onFetchFromCloud
}) => {
  const [cloudStats, setCloudStats] = useState<CloudStats | null>(null);
  const [isLoadingStats, setIsLoadingStats] = useState(false);
  const [isSyncingAction, setIsSyncingAction] = useState(false);
  const [activeView, setActiveView] = useState<'grid' | 'devices'>('grid');

  // Compute pending vs synced records
  const pendingRecords = useMemo(() => {
    return records.filter(r => r.syncStatus === 'pending_sync' || r.syncStatus === 'upload_failed');
  }, [records]);

  const loadStats = async () => {
    if (!isOnline) return;
    setIsLoadingStats(true);
    try {
      const stats = await fetchCloudCollectionStats();
      if (stats) setCloudStats(stats);
    } catch {
      // quiet catch
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
    if (isSyncingAction || !isOnline) return;
    setIsSyncingAction(true);
    try {
      await onManualSync();
      await loadStats();
    } finally {
      setIsSyncingAction(false);
    }
  };

  const handleTriggerDownload = async () => {
    if (isSyncingAction || !isOnline) return;
    setIsSyncingAction(true);
    try {
      await onFetchFromCloud(true, false);
      await loadStats();
    } finally {
      setIsSyncingAction(false);
    }
  };

  const getDeviceIcon = (devName?: string) => {
    const n = (devName || '').toLowerCase();
    if (n.includes('phone') || n.includes('iphone') || n.includes('android') || n.includes('mobile')) {
      return <Smartphone size={13} className="text-emerald-400" />;
    }
    if (n.includes('pad') || n.includes('tablet')) {
      return <Tablet size={13} className="text-cyan-400" />;
    }
    return <Laptop size={13} className="text-indigo-400" />;
  };

  // 6 Compact Data Modules
  const modules = useMemo(() => [
    {
      id: 'records',
      title: 'FFE & FH Logs',
      icon: Plane,
      accent: 'text-indigo-400 bg-indigo-500/10 border-indigo-500/20',
      local: records.length,
      cloud: cloudStats?.recordsCount ?? '-',
      pending: pendingRecords.length,
      synced: isOnline && cloudStats ? records.length === cloudStats.recordsCount && pendingRecords.length === 0 : isCloudSynced
    },
    {
      id: 'tempRecords',
      title: 'Movement (M)',
      icon: Truck,
      accent: 'text-emerald-400 bg-emerald-500/10 border-emerald-500/20',
      local: tempRecords.length,
      cloud: cloudStats?.tempRecordsCount ?? '-',
      pending: 0,
      synced: isOnline && cloudStats ? tempRecords.length === cloudStats.tempRecordsCount : isCloudSynced
    },
    {
      id: 'masterData',
      title: 'Master DB (MD)',
      icon: Database,
      accent: 'text-blue-400 bg-blue-500/10 border-blue-500/20',
      local: masterData.length,
      cloud: cloudStats?.masterDataCount ?? '-',
      pending: 0,
      synced: isOnline && cloudStats ? masterData.length === cloudStats.masterDataCount : isCloudSynced
    },
    {
      id: 'vehicleSummaries',
      title: 'Table Output (TO)',
      icon: Layers,
      accent: 'text-amber-400 bg-amber-500/10 border-amber-500/20',
      local: vehicleSummaries.length,
      cloud: cloudStats?.vehicleSummariesCount ?? '-',
      pending: 0,
      synced: isOnline && cloudStats ? vehicleSummaries.length === cloudStats.vehicleSummariesCount : isCloudSynced
    },
    {
      id: 'dossierHistory',
      title: 'Dossier (INV)',
      icon: FileSpreadsheet,
      accent: 'text-purple-400 bg-purple-500/10 border-purple-500/20',
      local: dossierHistory.length,
      cloud: cloudStats?.dossierHistoryCount ?? '-',
      pending: 0,
      synced: isOnline && cloudStats ? dossierHistory.length === cloudStats.dossierHistoryCount : isCloudSynced
    },
    {
      id: 'watchList',
      title: 'Watch-list (WL)',
      icon: ShieldAlert,
      accent: 'text-rose-400 bg-rose-500/10 border-rose-500/20',
      local: watchList.length,
      cloud: cloudStats?.watchListCount ?? '-',
      pending: 0,
      synced: isOnline && cloudStats ? watchList.length === cloudStats.watchListCount : isCloudSynced
    }
  ], [records, tempRecords, masterData, vehicleSummaries, dossierHistory, watchList, cloudStats, pendingRecords, isOnline, isCloudSynced]);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-slate-950/80 backdrop-blur-xs">
      <motion.div
        initial={{ opacity: 0, scale: 0.96, y: 8 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.96, y: 8 }}
        className="bg-slate-900 border border-slate-700/80 rounded-2xl w-full max-w-lg shadow-2xl overflow-hidden text-slate-100 flex flex-col"
      >
        {/* Compact Clean Header */}
        <div className="px-4 py-3 bg-slate-900 border-b border-slate-800 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="w-7 h-7 rounded-lg bg-indigo-600/30 text-indigo-400 flex items-center justify-center border border-indigo-500/30">
              <Server size={15} />
            </div>
            <div>
              <div className="flex items-center gap-1.5">
                <span className="text-xs font-black uppercase tracking-wider text-white">
                  Cloud Sync Monitor
                </span>
                <span className={`w-2 h-2 rounded-full ${isOnline ? 'bg-emerald-400 animate-pulse' : 'bg-rose-500'}`} />
              </div>
              <span className="text-[10px] text-slate-400 font-medium">
                {cloudAuthUser?.username || 'OFFICER'} • {cloudAuthUser?.deviceName || 'Device'}
              </span>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={() => setActiveView(activeView === 'grid' ? 'devices' : 'grid')}
              className="text-[10px] font-bold px-2 py-1 rounded-md bg-slate-800 hover:bg-slate-700 text-slate-300 border border-slate-700 transition-colors flex items-center gap-1 cursor-pointer"
            >
              <Smartphone size={11} />
              <span>{activeView === 'grid' ? `Devices (${deviceSessions.length})` : 'Show Grid'}</span>
            </button>
            <button
              onClick={onClose}
              className="w-7 h-7 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white flex items-center justify-center transition-colors cursor-pointer"
            >
              <X size={14} />
            </button>
          </div>
        </div>

        {/* Action Toolbar */}
        <div className="p-3 bg-slate-800/40 border-b border-slate-800 flex items-center justify-between gap-2">
          <div className="flex items-center gap-1.5 text-[10px] text-slate-400 font-mono">
            <span>Sync:</span>
            <span className="font-bold text-indigo-300">{lastSyncTime || 'Active'}</span>
            <button
              onClick={loadStats}
              disabled={isLoadingStats || !isOnline}
              title="Refresh Cloud Stats"
              className="text-indigo-400 hover:text-indigo-300 p-0.5 rounded transition-colors cursor-pointer"
            >
              <RefreshCw size={10} className={isLoadingStats ? "animate-spin" : ""} />
            </button>
          </div>

          <div className="flex items-center gap-1.5">
            <button
              onClick={handleTriggerUpload}
              disabled={isSyncingAction || !isOnline}
              className={`py-1.5 px-2.5 rounded-lg font-black uppercase text-[10px] flex items-center gap-1 shadow-xs transition-all ${
                !isOnline
                  ? 'bg-slate-800 text-slate-500 cursor-not-allowed border border-slate-700'
                  : isSyncingAction
                  ? 'bg-indigo-700 text-white cursor-wait'
                  : 'bg-indigo-600 hover:bg-indigo-500 text-white cursor-pointer'
              }`}
            >
              <UploadCloud size={12} className={isSyncingAction ? "animate-bounce" : ""} />
              <span>Upload Local</span>
            </button>

            <button
              onClick={handleTriggerDownload}
              disabled={isSyncingAction || !isOnline}
              className={`py-1.5 px-2.5 rounded-lg font-black uppercase text-[10px] flex items-center gap-1 border transition-all ${
                !isOnline
                  ? 'bg-slate-800 text-slate-600 border-slate-700 cursor-not-allowed'
                  : 'bg-slate-800 hover:bg-slate-700 text-cyan-300 border-slate-700 hover:border-cyan-500/40 cursor-pointer'
              }`}
            >
              <DownloadCloud size={12} />
              <span>Pull Cloud</span>
            </button>
          </div>
        </div>

        {/* Body Content */}
        <div className="p-3">
          {activeView === 'grid' ? (
            /* Compact 3-Column / 2-Row Matrix */
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
              {modules.map((m) => {
                const Icon = m.icon;
                return (
                  <div
                    key={m.id}
                    className="bg-slate-800/60 border border-slate-700/60 rounded-xl p-2.5 flex flex-col justify-between hover:border-slate-600 transition-colors"
                  >
                    <div className="flex items-center justify-between mb-1.5">
                      <div className="flex items-center gap-1.5">
                        <div className={`p-1 rounded-md border ${m.accent}`}>
                          <Icon size={12} />
                        </div>
                        <span className="text-[11px] font-bold text-slate-200 truncate">
                          {m.title}
                        </span>
                      </div>
                    </div>

                    <div className="flex items-baseline justify-between font-mono my-1 px-0.5">
                      <div className="text-left">
                        <span className="text-[9px] text-slate-400 block uppercase leading-none mb-0.5">Local</span>
                        <span className="text-sm font-black text-white">{m.local}</span>
                      </div>
                      <div className="text-slate-500 text-xs">/</div>
                      <div className="text-right">
                        <span className="text-[9px] text-cyan-400 block uppercase leading-none mb-0.5">Cloud</span>
                        <span className="text-sm font-black text-cyan-300">{m.cloud}</span>
                      </div>
                    </div>

                    <div className="mt-1 pt-1 border-t border-slate-700/40 flex items-center justify-between">
                      {m.pending > 0 ? (
                        <span className="text-[9px] font-bold text-amber-300 bg-amber-500/10 border border-amber-500/20 px-1.5 py-0.2 rounded flex items-center gap-0.5">
                          <Clock size={8} /> {m.pending} Pending
                        </span>
                      ) : m.synced ? (
                        <span className="text-[9px] font-bold text-emerald-300 bg-emerald-500/10 border border-emerald-500/20 px-1.5 py-0.2 rounded flex items-center gap-0.5">
                          <CheckCircle2 size={8} /> Synced
                        </span>
                      ) : (
                        <span className="text-[9px] font-bold text-slate-400 bg-slate-700/40 px-1.5 py-0.2 rounded flex items-center gap-0.5">
                          <Check size={8} /> Ready
                        </span>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          ) : (
            /* Connected Devices View */
            <div className="space-y-1.5 max-h-56 overflow-y-auto pr-1">
              {deviceSessions.length === 0 ? (
                <div className="p-4 text-center text-slate-400 text-xs">No secondary devices connected</div>
              ) : (
                deviceSessions.map((dev: any, i: number) => (
                  <div key={i} className="bg-slate-800/60 border border-slate-700/60 p-2 rounded-xl flex items-center justify-between text-xs">
                    <div className="flex items-center gap-2">
                      <div className="p-1.5 rounded-lg bg-slate-700/60 text-slate-300">
                        {getDeviceIcon(dev.deviceName)}
                      </div>
                      <div>
                        <div className="font-bold text-white text-xs">{dev.deviceName || 'Device'}</div>
                        <span className="text-[9px] text-slate-400 font-mono">User: {dev.username || 'OFFICER'}</span>
                      </div>
                    </div>
                    <span className="text-[9px] text-emerald-400 font-bold font-mono">● Connected</span>
                  </div>
                ))
              )}
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="px-3 py-2 bg-slate-900 border-t border-slate-800 text-center text-[9px] text-slate-500 uppercase tracking-wider font-mono">
          2-Way Auto-Sync Active • Myeik Immigration
        </div>
      </motion.div>
    </div>
  );
};
