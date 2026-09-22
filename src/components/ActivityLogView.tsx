import React, { useState, useEffect, useMemo } from 'react';
import { 
  ShieldAlert, Activity, Search, Filter, Download, Trash2, 
  RefreshCw, Calendar, User, Clock, ArrowUpDown, ChevronDown, 
  ChevronUp, Eye, CheckCircle, AlertTriangle, Database, FileSpreadsheet, Users
} from 'lucide-react';
import { ActivityLogEntry, ActivityActionType, ActivityModule } from '../types';
import { 
  getActivityLogs, 
  clearActivityLogs, 
  exportActivityLogsToExcel,
  subscribeToCloudActivityLogs 
} from '../utils/activityLogger';

interface ActivityLogViewProps {
  cloudAuthUser: { username: string; role: string } | null;
  showToast: (msg: string) => void;
  records?: any[];
  watchList?: any[];
  masterData?: any[];
  vehicleSummaries?: any[];
  checkingHistory?: any[];
}

export const ActivityLogView: React.FC<ActivityLogViewProps> = ({
  cloudAuthUser,
  showToast,
  records = [],
  watchList = [],
  masterData = [],
  vehicleSummaries = [],
  checkingHistory = []
}) => {
  const isSuperadmin = cloudAuthUser?.role === 'Superadmin';
  const [logs, setLogs] = useState<ActivityLogEntry[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [filterAction, setFilterAction] = useState<string>('ALL');
  const [filterModule, setFilterModule] = useState<string>('ALL');
  const [filterDate, setFilterDate] = useState<string>('');
  const [showClearConfirm, setShowClearConfirm] = useState<boolean>(false);
  const [expandedLogId, setExpandedLogId] = useState<string | null>(null);

  // System Wide Data Summary Breakdown
  const systemSummary = useMemo(() => {
    const ffeCount = records.filter(r => r.logType === 'FFE' || !r.logType).length;
    const fcrCount = records.filter(r => r.logType === 'FCR').length;
    const inCount = records.filter(r => r.mode === 'IN').length;
    const outCount = records.filter(r => r.mode === 'OUT').length;
    const totalRecords = records.length;
    const totalWatchList = watchList.length;
    const totalMaster = masterData.length;
    const totalVehicles = vehicleSummaries.length;
    const totalChecks = checkingHistory.length;

    return {
      totalRecords,
      ffeCount,
      fcrCount,
      inCount,
      outCount,
      totalWatchList,
      totalMaster,
      totalVehicles,
      totalChecks
    };
  }, [records, watchList, masterData, vehicleSummaries, checkingHistory]);

  const fetchLogs = async (forceCloud = true) => {
    setIsLoading(true);
    try {
      const data = await getActivityLogs(forceCloud);
      setLogs(data);
    } catch (e) {
      console.error(e);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    fetchLogs(true);
    
    // Subscribe to real-time cloud updates across all user actions
    const unsubscribeCloud = subscribeToCloudActivityLogs((cloudLogs) => {
      if (Array.isArray(cloudLogs) && cloudLogs.length > 0) {
        setLogs(cloudLogs);
      }
    });

    const handleLogEvent = () => {
      fetchLogs(false);
    };
    window.addEventListener('imm_activity_logged', handleLogEvent);
    
    return () => {
      if (unsubscribeCloud) unsubscribeCloud();
      window.removeEventListener('imm_activity_logged', handleLogEvent);
    };
  }, []);

  const filteredLogs = useMemo(() => {
    return logs.filter(log => {
      if (filterAction !== 'ALL' && log.action !== filterAction) return false;
      if (filterModule !== 'ALL' && log.module !== filterModule) return false;
      if (filterDate && !log.timestamp.startsWith(filterDate)) return false;

      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const officer = (log.officerName || '').toLowerCase();
        const details = (log.details || '').toLowerCase();
        const target = (log.targetId || '').toLowerCase();
        const device = (log.deviceId || '').toLowerCase();
        const role = (log.officerRole || '').toLowerCase();
        const time = (log.readableTime || '').toLowerCase();
        return (
          officer.includes(q) ||
          details.includes(q) ||
          target.includes(q) ||
          device.includes(q) ||
          role.includes(q) ||
          time.includes(q)
        );
      }
      return true;
    });
  }, [logs, filterAction, filterModule, filterDate, searchQuery]);

  // Statistics
  const stats = useMemo(() => {
    const todayStr = new Date().toISOString().split('T')[0];
    const todayLogs = logs.filter(l => l.timestamp.startsWith(todayStr));
    const creates = logs.filter(l => l.action === 'CREATE').length;
    const updates = logs.filter(l => l.action === 'UPDATE').length;
    const deletes = logs.filter(l => l.action === 'DELETE' || l.action === 'SESSION_KICK').length;

    return {
      total: logs.length,
      today: todayLogs.length,
      creates,
      updates,
      deletes
    };
  }, [logs]);

  const handleClearAll = async () => {
    await clearActivityLogs();
    setLogs([]);
    setShowClearConfirm(false);
    showToast("Audit Activity Logs အားလုံး ရှင်းလင်းပြီးပါပြီ");
  };

  const handleExport = () => {
    if (filteredLogs.length === 0) {
      showToast("ထုတ်ယူရန် Log မရှိပါ");
      return;
    }
    exportActivityLogsToExcel(filteredLogs);
    showToast(`Activity Log (${filteredLogs.length}) စောင်အား Excel အဖြစ် ထုတ်ယူလိုက်ပါပြီ`);
  };

  const getActionBadge = (action: ActivityActionType) => {
    switch (action) {
      case 'CREATE':
        return <span className="px-2 py-0.5 rounded-md text-[10px] font-black bg-emerald-100 text-emerald-800 border border-emerald-300">CREATE</span>;
      case 'UPDATE':
        return <span className="px-2 py-0.5 rounded-md text-[10px] font-black bg-blue-100 text-blue-800 border border-blue-300">UPDATE</span>;
      case 'DELETE':
        return <span className="px-2 py-0.5 rounded-md text-[10px] font-black bg-rose-100 text-rose-800 border border-rose-300 animate-pulse">DELETE</span>;
      case 'BACKUP':
        return <span className="px-2 py-0.5 rounded-md text-[10px] font-black bg-amber-100 text-amber-900 border border-amber-300">BACKUP</span>;
      case 'RESTORE':
        return <span className="px-2 py-0.5 rounded-md text-[10px] font-black bg-purple-100 text-purple-900 border border-purple-300">RESTORE</span>;
      case 'SYNC':
        return <span className="px-2 py-0.5 rounded-md text-[10px] font-black bg-cyan-100 text-cyan-900 border border-cyan-300">SYNC</span>;
      case 'WATCHLIST_MATCH':
        return <span className="px-2 py-0.5 rounded-md text-[10px] font-black bg-rose-600 text-white border border-rose-700">MATCH</span>;
      case 'SESSION_KICK':
        return <span className="px-2 py-0.5 rounded-md text-[10px] font-black bg-orange-100 text-orange-900 border border-orange-300">KICK</span>;
      default:
        return <span className="px-2 py-0.5 rounded-md text-[10px] font-black bg-slate-100 text-slate-800 border border-slate-300">{action}</span>;
    }
  };

  const getModuleBadge = (module: ActivityModule) => {
    switch (module) {
      case 'FFE':
        return <span className="px-2 py-0.5 rounded bg-indigo-50 text-indigo-700 font-extrabold text-[10px]">FFE Entry</span>;
      case 'FCR':
        return <span className="px-2 py-0.5 rounded bg-teal-50 text-teal-700 font-extrabold text-[10px]">Form C</span>;
      case 'CHECKING':
        return <span className="px-2 py-0.5 rounded bg-sky-50 text-sky-700 font-extrabold text-[10px]">Checking</span>;
      case 'MASTER':
        return <span className="px-2 py-0.5 rounded bg-amber-50 text-amber-700 font-extrabold text-[10px]">Master Data</span>;
      case 'WATCHLIST':
        return <span className="px-2 py-0.5 rounded bg-rose-50 text-rose-700 font-extrabold text-[10px]">Watchlist</span>;
      case 'SYSTEM':
        return <span className="px-2 py-0.5 rounded bg-slate-100 text-slate-700 font-extrabold text-[10px]">System</span>;
      case 'CRT':
        return <span className="px-2 py-0.5 rounded bg-purple-50 text-purple-700 font-extrabold text-[10px]">CRT Report</span>;
      case 'TELEG':
        return <span className="px-2 py-0.5 rounded bg-blue-50 text-blue-700 font-extrabold text-[10px]">Telegraph</span>;
      case 'CLOUD_AUTH':
        return <span className="px-2 py-0.5 rounded bg-emerald-50 text-emerald-700 font-extrabold text-[10px]">Auth</span>;
      default:
        return <span className="px-2 py-0.5 rounded bg-slate-50 text-slate-600 font-extrabold text-[10px]">{module}</span>;
    }
  };

  if (!isSuperadmin) {
    return (
      <div className="bg-white rounded-3xl p-8 border border-slate-200 shadow-xl text-center space-y-4 max-w-xl mx-auto my-8">
        <div className="w-16 h-16 bg-purple-100 text-purple-700 rounded-full flex items-center justify-center mx-auto shadow-inner">
          <ShieldAlert size={32} />
        </div>
        <h3 className="text-xl font-black text-slate-900 uppercase tracking-tight">Restricted Access (Superadmin Only)</h3>
        <p className="text-sm text-slate-600 font-medium leading-relaxed">
          ဤ Activity & Audit Log ကဏ္ဍသည် စနစ်တွင်း မှတ်တမ်းပြင်ဆင်မှု၊ ဖျက်ပစ်မှုနှင့် လုပ်ဆောင်ချက်များအား ခြေရာခံစစ်ဆေးရန် <b>SUPERADMIN</b> အဆင့် အကောင့်များအတွက်သာ သီးသန့် ဖွင့်လှစ်ထားရှိပါသည်။
        </p>
        <div className="pt-2">
          <span className="inline-block px-3 py-1 bg-amber-50 text-amber-800 text-xs font-bold rounded-xl border border-amber-200">
            လက်ရှိအကောင့် အဆင့်အတန်း: {cloudAuthUser?.role || 'Guest / Editor'}
          </span>
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-6">
      {/* Superadmin Header Deck */}
      <div className="bg-gradient-to-r from-purple-950 via-slate-900 to-indigo-950 text-white p-6 sm:p-8 rounded-3xl shadow-xl border border-purple-800/40">
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
          <div className="flex items-center gap-3">
            <div className="p-3 bg-purple-600/30 text-purple-300 rounded-2xl border border-purple-400/30">
              <Activity size={28} />
            </div>
            <div>
              <div className="flex items-center gap-2 flex-wrap">
                <h2 className="text-xl sm:text-2xl font-black uppercase tracking-tight">Activity & Audit Log</h2>
                <span className="px-2.5 py-0.5 rounded-full text-[10px] font-black bg-purple-500 text-white uppercase tracking-wider">
                  👑 Superadmin Only
                </span>
                <span className="px-2.5 py-0.5 rounded-full text-[10px] font-bold bg-emerald-500/30 text-emerald-300 border border-emerald-400/30 uppercase tracking-wider">
                  👥 All Users Logged (Max 300 Records)
                </span>
              </div>
              <p className="text-xs text-purple-200/80 font-medium mt-1">
                အသုံးပြုသူအားလုံး (Superadmin, Editor, Viewer, Officer) ၏ မှတ်တမ်းဖြည့်သွင်းခြင်း၊ ပြင်ဆင်ခြင်း၊ ဖျက်ခြင်းနှင့် စနစ်ထိန်းချုပ်မှု လုပ်ဆောင်ချက်များ ခြေရာခံခြင်း
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2 self-stretch sm:self-auto flex-wrap">
            <button
              onClick={() => fetchLogs(true)}
              className="px-3.5 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-purple-200 text-xs font-bold transition-all flex items-center gap-1.5 cursor-pointer shadow-sm"
              title="Refresh Logs"
            >
              <RefreshCw size={14} className={isLoading ? "animate-spin" : ""} />
              <span>Refresh</span>
            </button>
            <button
              onClick={handleExport}
              className="px-3.5 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold transition-all flex items-center gap-1.5 cursor-pointer shadow-sm"
              title="Excel အဖြစ် သိမ်းမည်"
            >
              <FileSpreadsheet size={14} />
              <span>Export Excel</span>
            </button>
            {isSuperadmin && (
              <button
                onClick={() => setShowClearConfirm(true)}
                className="px-3.5 py-2 rounded-xl bg-rose-600/80 hover:bg-rose-700 text-white text-xs font-bold transition-all flex items-center gap-1.5 cursor-pointer shadow-sm"
                title="Logs များ ရှင်းလင်းမည်"
              >
                <Trash2 size={14} />
                <span>Clear Logs</span>
              </button>
            )}
          </div>
        </div>

        {/* Overall System Data Inventory Summary (ဒေတာ အားလုံး မည်မျှရှိကြောင်း အနှစ်ချုပ်) */}
        <div className="mt-6 pt-6 border-t border-purple-800/40 space-y-3">
          <div className="flex items-center justify-between">
            <span className="text-xs font-black uppercase text-purple-300 tracking-wider flex items-center gap-1.5">
              <Database size={14} className="text-purple-400" />
              စနစ်တွင်း ဒေတာ အားလုံး စာရင်း အနှစ်ချုပ် (System Data Overview)
            </span>
            <span className="text-[10px] font-bold text-purple-200/70">
              Live Total Records
            </span>
          </div>

          <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-6 gap-2.5">
            <div className="bg-slate-900/80 border border-purple-500/30 rounded-2xl p-3 text-center shadow-inner">
              <div className="text-[10px] text-slate-400 font-black uppercase">စုစုပေါင်း လူဦးရေ/Flight</div>
              <div className="text-lg sm:text-xl font-black text-amber-300">{systemSummary.totalRecords}</div>
              <div className="text-[9px] text-slate-400 mt-0.5">Total Imm Records</div>
            </div>

            <div className="bg-slate-900/80 border border-purple-500/30 rounded-2xl p-3 text-center shadow-inner">
              <div className="text-[10px] text-slate-400 font-black uppercase">FFE လေယာဉ် / FCR</div>
              <div className="text-lg sm:text-xl font-black text-cyan-300">
                {systemSummary.ffeCount} <span className="text-xs text-slate-400 font-normal">/ {systemSummary.fcrCount}</span>
              </div>
              <div className="text-[9px] text-slate-400 mt-0.5">Flight vs Form C</div>
            </div>

            <div className="bg-slate-900/80 border border-purple-500/30 rounded-2xl p-3 text-center shadow-inner">
              <div className="text-[10px] text-slate-400 font-black uppercase">ဝင်ရောက် (IN) / ထွက်ခွာ (OUT)</div>
              <div className="text-lg sm:text-xl font-black text-emerald-300">
                {systemSummary.inCount} <span className="text-xs text-slate-400 font-normal">/ {systemSummary.outCount}</span>
              </div>
              <div className="text-[9px] text-slate-400 mt-0.5">Inbound / Outbound</div>
            </div>

            <div className="bg-slate-900/80 border border-purple-500/30 rounded-2xl p-3 text-center shadow-inner">
              <div className="text-[10px] text-slate-400 font-black uppercase">စောင့်ကြည့်စာရင်း (WL)</div>
              <div className="text-lg sm:text-xl font-black text-rose-300">{systemSummary.totalWatchList}</div>
              <div className="text-[9px] text-slate-400 mt-0.5">Watch-list Persons</div>
            </div>

            <div className="bg-slate-900/80 border border-purple-500/30 rounded-2xl p-3 text-center shadow-inner">
              <div className="text-[10px] text-slate-400 font-black uppercase">အခြေခံ Master DB</div>
              <div className="text-lg sm:text-xl font-black text-indigo-300">{systemSummary.totalMaster}</div>
              <div className="text-[9px] text-slate-400 mt-0.5">Options / Codes</div>
            </div>

            <div className="bg-slate-900/80 border border-purple-500/30 rounded-2xl p-3 text-center shadow-inner">
              <div className="text-[10px] text-slate-400 font-black uppercase">ယာဉ်ဇယား (Vehicles)</div>
              <div className="text-lg sm:text-xl font-black text-teal-300">{systemSummary.totalVehicles}</div>
              <div className="text-[9px] text-slate-400 mt-0.5">Vehicle Summaries</div>
            </div>
          </div>
        </div>

        {/* Quick Audit Action Stats Banner */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 pt-6 mt-6 border-t border-purple-800/40">
          <div className="bg-purple-900/40 border border-purple-700/40 rounded-2xl p-3 text-center">
            <div className="text-[10px] text-purple-300 font-black uppercase">Total Audit Logs</div>
            <div className="text-lg sm:text-xl font-black text-white">{stats.total}</div>
          </div>
          <div className="bg-purple-900/40 border border-purple-700/40 rounded-2xl p-3 text-center">
            <div className="text-[10px] text-purple-300 font-black uppercase">Today Actions</div>
            <div className="text-lg sm:text-xl font-black text-emerald-400">{stats.today}</div>
          </div>
          <div className="bg-purple-900/40 border border-purple-700/40 rounded-2xl p-3 text-center">
            <div className="text-[10px] text-purple-300 font-black uppercase">Created / Updated</div>
            <div className="text-lg sm:text-xl font-black text-blue-300">{stats.creates + stats.updates}</div>
          </div>
          <div className="bg-purple-900/40 border border-purple-700/40 rounded-2xl p-3 text-center">
            <div className="text-[10px] text-purple-300 font-black uppercase">Deleted / Revoked</div>
            <div className="text-lg sm:text-xl font-black text-rose-300">{stats.deletes}</div>
          </div>
        </div>
      </div>

      {/* Filter and Search Bar */}
      <div className="bg-white rounded-2xl p-4 border border-slate-200 shadow-sm space-y-3">
        <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3">
          {/* Search */}
          <div className="relative">
            <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="အရာရှိ၊ ပတ်စပို့၊ အကြောင်းအရာ..."
              className="w-full pl-9 pr-3 py-2 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-purple-500/20 focus:border-purple-500 outline-none font-medium"
            />
          </div>

          {/* Action Filter */}
          <div className="flex items-center gap-1.5">
            <span className="text-[11px] font-bold text-slate-500 shrink-0">Action:</span>
            <select
              value={filterAction}
              onChange={(e) => setFilterAction(e.target.value)}
              className="w-full py-2 px-2.5 text-xs bg-slate-50 border border-slate-200 rounded-xl font-bold text-slate-700 outline-none cursor-pointer"
            >
              <option value="ALL">All Actions</option>
              <option value="CREATE">CREATE (အသစ်ထည့်)</option>
              <option value="UPDATE">UPDATE (ပြင်ဆင်)</option>
              <option value="DELETE">DELETE (ဖျက်ပစ်)</option>
              <option value="BACKUP">BACKUP (အရံသိမ်း)</option>
              <option value="RESTORE">RESTORE (ပြန်ယူ)</option>
              <option value="SYNC">SYNC (Cloud Sync)</option>
              <option value="WATCHLIST_MATCH">WATCHLIST (တိုက်ဆိုင်တွေ့)</option>
              <option value="SESSION_KICK">SESSION_KICK (အကောင့်ထုတ်)</option>
            </select>
          </div>

          {/* Module Filter */}
          <div className="flex items-center gap-1.5">
            <span className="text-[11px] font-bold text-slate-500 shrink-0">Module:</span>
            <select
              value={filterModule}
              onChange={(e) => setFilterModule(e.target.value)}
              className="w-full py-2 px-2.5 text-xs bg-slate-50 border border-slate-200 rounded-xl font-bold text-slate-700 outline-none cursor-pointer"
            >
              <option value="ALL">All Modules</option>
              <option value="FFE">FFE Flight Entry</option>
              <option value="FCR">Form C</option>
              <option value="CHECKING">Checking / Still In</option>
              <option value="MASTER">Master Data (MD)</option>
              <option value="WATCHLIST">Watchlist</option>
              <option value="SYSTEM">System Settings</option>
              <option value="CRT">Custom Report Table</option>
              <option value="TELEG">Telegraph (T&T)</option>
            </select>
          </div>

          {/* Date Filter */}
          <div className="flex items-center gap-1.5">
            <span className="text-[11px] font-bold text-slate-500 shrink-0">Date:</span>
            <input
              type="date"
              value={filterDate}
              onChange={(e) => setFilterDate(e.target.value)}
              className="w-full py-2 px-2.5 text-xs bg-slate-50 border border-slate-200 rounded-xl font-bold text-slate-700 outline-none"
            />
            {filterDate && (
              <button
                type="button"
                onClick={() => setFilterDate('')}
                className="text-xs text-rose-500 font-black px-1.5 hover:underline"
              >
                Clear
              </button>
            )}
          </div>
        </div>
      </div>

      {/* Logs Table Card */}
      <div className="bg-white rounded-3xl border border-slate-200 shadow-md overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs border-collapse">
            <thead>
              <tr className="bg-slate-900 text-white font-black uppercase text-[10px] tracking-wider border-b border-slate-800">
                <th className="p-3 text-center w-12">စဉ်</th>
                <th className="p-3 text-left w-36">အချိန် / ရက်စွဲ</th>
                <th className="p-3 text-center w-24">လုပ်ဆောင်ချက်</th>
                <th className="p-3 text-center w-24">ကဏ္ဍ</th>
                <th className="p-3 text-left w-36">အရာရှိ / Role</th>
                <th className="p-3 text-left w-32">ပတ်စပို့ / အညွှန်း</th>
                <th className="p-3 text-left">အသေးစိတ်မှတ်တမ်း (Details)</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 font-medium">
              {filteredLogs.length === 0 ? (
                <tr>
                  <td colSpan={7} className="p-8 text-center text-slate-400 italic">
                    {isLoading ? "Logs များအား ရယူနေပါသည်..." : "မှတ်တမ်းတင်ထားသော Activity မရှိသေးပါ"}
                  </td>
                </tr>
              ) : (
                filteredLogs.map((log, idx) => {
                  const isExpanded = expandedLogId === log.id;
                  return (
                    <React.Fragment key={log.id}>
                      <tr 
                        className={`hover:bg-slate-50/80 transition-colors ${
                          log.action === 'DELETE' ? 'bg-rose-50/30' : ''
                        }`}
                      >
                        <td className="p-3 text-center text-slate-400 font-mono font-bold">
                          {idx + 1}
                        </td>
                        <td className="p-3 font-mono text-[11px] text-slate-600 whitespace-nowrap">
                          {log.readableTime}
                        </td>
                        <td className="p-3 text-center">
                          {getActionBadge(log.action)}
                        </td>
                        <td className="p-3 text-center">
                          {getModuleBadge(log.module)}
                        </td>
                        <td className="p-3">
                          <div className="font-extrabold text-slate-900 text-[11px] truncate max-w-[140px]" title={log.officerName}>
                            {log.officerName}
                          </div>
                          <div className="text-[10px] text-slate-400 font-bold uppercase">
                            {log.officerRole}
                          </div>
                        </td>
                        <td className="p-3">
                          {log.targetId ? (
                            <span className="font-mono font-black text-indigo-900 bg-indigo-50 px-2 py-0.5 rounded border border-indigo-200">
                              {log.targetId}
                            </span>
                          ) : (
                            <span className="text-slate-400">-</span>
                          )}
                        </td>
                        <td className="p-3">
                          <div className="flex items-center justify-between gap-2">
                            <span className="text-slate-700 font-semibold line-clamp-1">
                              {log.details}
                            </span>
                            <button
                              onClick={() => setExpandedLogId(isExpanded ? null : log.id)}
                              className="text-slate-400 hover:text-indigo-600 p-1 shrink-0 cursor-pointer"
                              title="အသေးစိတ်ကြည့်မည်"
                            >
                              {isExpanded ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
                            </button>
                          </div>
                        </td>
                      </tr>
                      {/* Expanded Details Row */}
                      {isExpanded && (
                        <tr className="bg-slate-50/90 border-y border-slate-200 text-xs">
                          <td colSpan={7} className="p-4 space-y-2">
                            <div className="grid grid-cols-1 md:grid-cols-3 gap-3 bg-white p-3.5 rounded-xl border border-slate-200 text-slate-700">
                              <div>
                                <span className="text-[10px] font-black uppercase text-slate-400 block">Log Entry ID:</span>
                                <span className="font-mono text-xs font-bold text-purple-900">{log.id}</span>
                              </div>
                              <div>
                                <span className="text-[10px] font-black uppercase text-slate-400 block">Device Identifier:</span>
                                <span className="font-mono text-xs font-semibold text-slate-700">{log.deviceId || 'local_terminal'}</span>
                              </div>
                              <div>
                                <span className="text-[10px] font-black uppercase text-slate-400 block">Exact ISO Timestamp:</span>
                                <span className="font-mono text-[11px] text-slate-600">{log.timestamp}</span>
                              </div>
                            </div>
                            <div className="bg-white p-3.5 rounded-xl border border-slate-200">
                              <span className="text-[10px] font-black uppercase text-slate-400 block mb-1">Full Detailed Action:</span>
                              <p className="text-xs text-slate-800 font-medium leading-relaxed font-pyidaungsu whitespace-pre-wrap">
                                {log.details}
                              </p>
                            </div>
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Clear Confirmation Modal */}
      {showClearConfirm && (
        <div className="fixed inset-0 z-[600] flex items-center justify-center p-4 bg-slate-900/80 backdrop-blur-sm">
          <div className="bg-white rounded-3xl p-6 sm:p-8 max-w-md w-full shadow-2xl border-t-8 border-rose-600 space-y-5 text-center">
            <div className="w-16 h-16 bg-rose-100 text-rose-600 rounded-full flex items-center justify-center mx-auto">
              <AlertTriangle size={32} />
            </div>
            <div className="space-y-2">
              <h3 className="text-xl font-black text-slate-900 uppercase">Clear All Activity Logs?</h3>
              <p className="text-xs text-slate-600 font-semibold leading-relaxed">
                ဤလုပ်ဆောင်ချက်သည် စနစ်အတွင်း မှတ်တမ်းတင်ထားသော Activity/Audit Log အားလုံးကို အပြီးတိုင် ဖျက်ပစ်ပါမည်။ ဤလုပ်ဆောင်ချက်အား ပြန်လည် မရယူနိုင်ပါ။
              </p>
            </div>
            <div className="flex gap-3 pt-2">
              <button
                type="button"
                onClick={() => setShowClearConfirm(false)}
                className="flex-1 py-3 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl text-xs transition-all cursor-pointer"
              >
                မလုပ်တော့ပါ (Cancel)
              </button>
              <button
                type="button"
                onClick={handleClearAll}
                className="flex-1 py-3 bg-rose-600 hover:bg-rose-700 text-white font-bold rounded-xl text-xs transition-all cursor-pointer shadow-lg shadow-rose-200"
              >
                သေချာသည် ဖျက်မည်
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
