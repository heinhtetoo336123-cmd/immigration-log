import React, { useState, useEffect } from 'react';
import { 
  Server, 
  UploadCloud, 
  DownloadCloud, 
  X, 
  CheckCircle2, 
  AlertCircle, 
  Clock, 
  Laptop, 
  Smartphone, 
  Tablet, 
  RefreshCw 
} from 'lucide-react';
import { motion } from 'framer-motion';
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
  isCloudSynced,
  lastSyncTime,
  isQuotaExhausted,
  cloudAuthUser,
  deviceSessions,
  records,
  onManualSync,
  onFetchFromCloud,
  onResetQuota
}) => {
  const [activeTab, setActiveTab] = useState<'overview' | 'pending' | 'devices'>('overview');
  const [cloudStats, setCloudStats] = useState<CloudStats | null>(null);
  const [isLoadingStats, setIsLoadingStats] = useState(false);
  const [isSyncingAction, setIsSyncingAction] = useState(false);
  const [statusMessage, setStatusMessage] = useState<string | null>(null);

  const pendingRecords = records.filter(r => r.syncStatus === 'pending_sync' || r.syncStatus === 'upload_failed');
  const syncedCount = records.length - pendingRecords.length;

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
      setStatusMessage("✓ Download completed successfully.");
    } catch (err: any) {
      setStatusMessage(`Download failed: ${err?.message || 'Server error'}`);
    } finally {
      setIsSyncingAction(false);
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
            className="text-slate-400 hover:text-white p-1 rounded-lg hover:bg-slate-800 transition-colors"
          >
            <X size={16} />
          </button>
        </div>

        {/* Minimalist Tab Bar */}
        <div className="flex border-b border-slate-800 bg-slate-900/80 px-4 pt-2 gap-2 text-xs font-semibold">
          {[
            { id: 'overview', label: 'Overview' },
            { id: 'pending', label: `Pending (${pendingRecords.length})` },
            { id: 'devices', label: `Devices (${deviceSessions.length})` }
          ].map(tab => (
            <button
              key={tab.id}
              onClick={() => setActiveTab(tab.id as any)}
              className={`pb-2 px-3 border-b-2 transition-all cursor-pointer ${
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
                  <span className="text-[10px] text-slate-400 uppercase font-semibold block">Pending</span>
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
                      className="px-2 py-1 bg-rose-900/60 hover:bg-rose-800 text-white rounded text-[10px] font-bold"
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
                  <div className="text-xs font-semibold text-slate-400">No pending records</div>
                  <div className="text-[11px]">All records are in sync with cloud storage.</div>
                </div>
              ) : (
                <div className="space-y-1.5 max-h-72 overflow-y-auto pr-1">
                  {pendingRecords.map((r, idx) => (
                    <div
                      key={r.id || idx}
                      className="bg-slate-800/40 border border-slate-800 p-2.5 rounded-xl flex items-center justify-between text-xs"
                    >
                      <div>
                        <div className="flex items-center gap-2">
                          <span className="font-mono font-bold text-indigo-300">{r.passport}</span>
                          <span className="text-[10px] text-slate-400">[{r.mode}]</span>
                          <span className="text-[10px] text-slate-300 font-semibold">{r.fullname}</span>
                        </div>
                        <div className="text-[10px] text-slate-500 font-mono mt-0.5">
                          {r.timestamp} • {r.nationality || '-'}
                        </div>
                      </div>
                      <span className="text-[10px] text-amber-400 font-semibold px-2 py-0.5 bg-amber-950/40 rounded border border-amber-800/40">
                        Pending
                      </span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {activeTab === 'devices' && (
            <div className="space-y-2 max-h-72 overflow-y-auto pr-1">
              {deviceSessions.length === 0 ? (
                <div className="py-8 text-center text-slate-500 text-xs">
                  No other active device sessions recorded.
                </div>
              ) : (
                deviceSessions.map((session, idx) => {
                  const isCurrent = session.deviceId === cloudAuthUser?.deviceId;
                  return (
                    <div
                      key={session.deviceId || idx}
                      className={`p-2.5 rounded-xl border flex items-center justify-between text-xs ${
                        isCurrent 
                          ? 'bg-indigo-950/20 border-indigo-800/50' 
                          : 'bg-slate-800/30 border-slate-800'
                      }`}
                    >
                      <div className="flex items-center gap-2.5">
                        {getDeviceIcon(session.deviceName)}
                        <div>
                          <div className="font-semibold text-slate-200 flex items-center gap-1.5">
                            <span>{session.deviceName || 'Device'}</span>
                            {isCurrent && (
                              <span className="text-[9px] bg-indigo-500/30 text-indigo-300 px-1 rounded font-bold">
                                Current
                              </span>
                            )}
                          </div>
                          <div className="text-[10px] text-slate-500 font-mono">
                            {session.username || 'User'} • {session.accountRole || 'Editor'}
                          </div>
                        </div>
                      </div>
                      <div className="text-right text-[10px] text-slate-500">
                        {session.lastActive || 'Active'}
                      </div>
                    </div>
                  );
                })
              )}
            </div>
          )}
        </div>

        {/* Minimalist Footer */}
        <div className="px-5 py-3 border-t border-slate-800 flex justify-end">
          <button
            onClick={onClose}
            className="px-4 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white rounded-lg text-xs font-semibold transition-colors"
          >
            Close
          </button>
        </div>
      </motion.div>
    </div>
  );
};
