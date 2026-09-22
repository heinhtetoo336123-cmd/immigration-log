import React, { useState } from 'react';
import { 
  Bookmark, Check, Plus, Trash2, X, Sparkles, Layers, 
  FileText, Shield, ArrowRight, CheckCircle2
} from 'lucide-react';
import { CRTPreset, CRTColumnId } from '../types';
import { 
  DEFAULT_BUILT_IN_PRESETS, 
  getStoredCustomPresets, 
  saveCustomPreset, 
  deleteCustomPreset 
} from '../utils/crtPresets';

interface CustomReportPresetsModalProps {
  isOpen: boolean;
  onClose: () => void;
  currentSettings: {
    tableTitle: string;
    showSigner?: boolean;
    officerTitle: string;
    officerName: string;
    useBurmeseDigits: boolean;
    dataPool: 'STILL_IN' | 'ALL_MOVEMENTS' | 'INBOUND' | 'OUTBOUND';
    visibleColumns: CRTColumnId[];
    filterNationalities?: string[];
    filterVisaTypes?: string[];
    filterAddresses?: string[];
    filterNationality?: string;
    filterVisaType?: string;
    filterAddress?: string;
    filterAddressType?: 'ALL' | 'COMPANY' | 'OTHER';
    filterGender?: 'ALL' | 'M' | 'F';
    minElapsedDays?: string;
  };
  onApplyPreset: (preset: CRTPreset) => void;
  showToast: (msg: string) => void;
}

export const CustomReportPresetsModal: React.FC<CustomReportPresetsModalProps> = ({
  isOpen,
  onClose,
  currentSettings,
  onApplyPreset,
  showToast
}) => {
  const [customPresets, setCustomPresets] = useState<CRTPreset[]>(getStoredCustomPresets());
  const [activeTab, setActiveTab] = useState<'list' | 'save'>('list');
  const [newPresetName, setNewPresetName] = useState<string>('');
  const [newPresetDesc, setNewPresetDesc] = useState<string>('');

  if (!isOpen) return null;

  const handleSavePreset = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newPresetName.trim()) {
      showToast("Template အမည် ထည့်သွင်းပေးပါ");
      return;
    }

    const saved = saveCustomPreset({
      name: newPresetName.trim(),
      description: newPresetDesc.trim() || undefined,
      tableTitle: currentSettings.tableTitle,
      showSigner: currentSettings.showSigner,
      officerTitle: currentSettings.officerTitle,
      officerName: currentSettings.officerName,
      useBurmeseDigits: currentSettings.useBurmeseDigits,
      dataPool: currentSettings.dataPool,
      visibleColumns: currentSettings.visibleColumns,
      filterNationalities: currentSettings.filterNationalities,
      filterVisaTypes: currentSettings.filterVisaTypes,
      filterAddresses: currentSettings.filterAddresses,
      filterNationality: currentSettings.filterNationality,
      filterVisaType: currentSettings.filterVisaType,
      filterAddress: currentSettings.filterAddress,
      filterAddressType: currentSettings.filterAddressType,
      filterGender: currentSettings.filterGender,
      minElapsedDays: currentSettings.minElapsedDays
    });

    setCustomPresets(getStoredCustomPresets());
    setNewPresetName('');
    setNewPresetDesc('');
    setActiveTab('list');
    showToast(`Template "${saved.name}" အား အောင်မြင်စွာ မှတ်ထားပြီးပါပြီ`);
  };

  const handleDelete = (id: string, name: string) => {
    deleteCustomPreset(id);
    setCustomPresets(getStoredCustomPresets());
    showToast(`Template "${name}" အား ဖျက်ပစ်ပြီးပါပြီ`);
  };

  const allPresets: CRTPreset[] = [...DEFAULT_BUILT_IN_PRESETS, ...customPresets];

  return (
    <div className="fixed inset-0 z-[600] flex items-center justify-center p-4 bg-slate-900/80 backdrop-blur-xs animate-in fade-in duration-150">
      <div className="bg-white rounded-3xl p-6 max-w-2xl w-full shadow-2xl border border-slate-200 space-y-5">
        {/* Modal Header */}
        <div className="flex items-center justify-between border-b border-slate-100 pb-3">
          <div className="flex items-center gap-2.5">
            <div className="p-2.5 bg-purple-100 text-purple-700 rounded-2xl">
              <Bookmark size={22} />
            </div>
            <div>
              <h3 className="text-lg font-black text-slate-900 uppercase">
                Report Template Presets
              </h3>
              <p className="text-xs text-slate-500 font-medium">
                ပုံမှန်ထုတ်နေကျ အစီရင်ခံစာပုံစံများအား တစ်ချက်နှိပ်ရုံဖြင့် အလွယ်တကူ အသုံးပြုနိုင်ပါသည်
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded-xl transition-colors cursor-pointer"
          >
            <X size={18} />
          </button>
        </div>

        {/* Tab switcher */}
        <div className="flex bg-slate-100 p-1 rounded-2xl border border-slate-200">
          <button
            type="button"
            onClick={() => setActiveTab('list')}
            className={`flex-1 py-2 rounded-xl text-xs font-black transition-all flex items-center justify-center gap-2 cursor-pointer ${
              activeTab === 'list'
                ? 'bg-white text-indigo-950 shadow-xs'
                : 'text-slate-600 hover:text-indigo-600'
            }`}
          >
            <Layers size={14} />
            <span>Templates စာရင်း ({allPresets.length})</span>
          </button>
          <button
            type="button"
            onClick={() => setActiveTab('save')}
            className={`flex-1 py-2 rounded-xl text-xs font-black transition-all flex items-center justify-center gap-2 cursor-pointer ${
              activeTab === 'save'
                ? 'bg-white text-indigo-950 shadow-xs'
                : 'text-slate-600 hover:text-indigo-600'
            }`}
          >
            <Plus size={14} />
            <span>လက်ရှိပုံစံအား Template အသစ်အဖြစ် သိမ်းမည်</span>
          </button>
        </div>

        {/* Content Area */}
        {activeTab === 'list' ? (
          <div className="space-y-3 max-h-[380px] overflow-y-auto pr-1">
            <div className="grid grid-cols-1 gap-2.5">
              {allPresets.map((preset) => (
                <div
                  key={preset.id}
                  className="bg-slate-50 hover:bg-indigo-50/40 p-4 rounded-2xl border border-slate-200 hover:border-indigo-300 transition-all flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 group"
                >
                  <div className="space-y-1 flex-1">
                    <div className="flex items-center gap-2">
                      <span className="font-extrabold text-sm text-slate-900 group-hover:text-indigo-950 font-pyidaungsu">
                        {preset.name}
                      </span>
                      {preset.isBuiltIn ? (
                        <span className="px-2 py-0.5 rounded-full text-[9px] font-black bg-indigo-100 text-indigo-800 border border-indigo-200">
                          စနစ်ပုံသေ (Built-in)
                        </span>
                      ) : (
                        <span className="px-2 py-0.5 rounded-full text-[9px] font-black bg-emerald-100 text-emerald-800 border border-emerald-200">
                          စိတ်ကြိုက် (Custom)
                        </span>
                      )}
                    </div>
                    {preset.description && (
                      <p className="text-[11px] text-slate-500 font-medium font-pyidaungsu line-clamp-1">
                        {preset.description}
                      </p>
                    )}
                    <div className="flex items-center gap-2 text-[10px] text-slate-500 font-bold flex-wrap pt-0.5">
                      <span className="bg-white px-2 py-0.5 rounded border border-slate-200">
                        Title: {preset.tableTitle.substring(0, 24)}...
                      </span>
                      <span className="bg-white px-2 py-0.5 rounded border border-slate-200">
                        Columns: {preset.visibleColumns.length} ခု
                      </span>
                      <span className="bg-white px-2 py-0.5 rounded border border-slate-200">
                        Pool: {preset.dataPool}
                      </span>
                      {preset.filterAddressType && preset.filterAddressType !== 'ALL' && (
                        <span className="bg-indigo-50 text-indigo-700 font-black px-2 py-0.5 rounded border border-indigo-200">
                          {preset.filterAddressType === 'COMPANY' ? '🏢 ကုမ္ပဏီ (Co.Ltd)' : '🏠 အခြားလိပ်စာ (Other)'}
                        </span>
                      )}
                      {preset.filterAddresses && preset.filterAddresses.length > 0 && (
                        <span className="bg-amber-50 text-amber-800 font-black px-2 py-0.5 rounded border border-amber-200">
                          📍 လိပ်စာ ({preset.filterAddresses.length}) ခု
                        </span>
                      )}
                      {preset.filterNationalities && preset.filterNationalities.length > 0 && (
                        <span className="bg-purple-50 text-purple-700 font-black px-2 py-0.5 rounded border border-purple-200">
                          🌐 နိုင်ငံ ({preset.filterNationalities.length}) ခု
                        </span>
                      )}
                      {preset.filterVisaTypes && preset.filterVisaTypes.length > 0 && (
                        <span className="bg-emerald-50 text-emerald-800 font-black px-2 py-0.5 rounded border border-emerald-200">
                          🎫 ဗီဇာ ({preset.filterVisaTypes.length}) ခု
                        </span>
                      )}
                    </div>
                  </div>

                  <div className="flex items-center gap-2 self-end sm:self-center shrink-0">
                    {!preset.isBuiltIn && (
                      <button
                        type="button"
                        onClick={() => handleDelete(preset.id, preset.name)}
                        className="p-2 text-rose-500 hover:bg-rose-100 rounded-xl transition-colors cursor-pointer"
                        title="Delete Template"
                      >
                        <Trash2 size={15} />
                      </button>
                    )}
                    <button
                      type="button"
                      onClick={() => {
                        onApplyPreset(preset);
                        onClose();
                        showToast(`Template "${preset.name}" အား အောင်မြင်စွာ ပြောင်းလဲလိုက်ပါပြီ`);
                      }}
                      className="px-3.5 py-2 bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-black rounded-xl shadow-xs transition-all flex items-center gap-1.5 cursor-pointer"
                    >
                      <span>အသုံးပြုမည်</span>
                      <ArrowRight size={13} />
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        ) : (
          /* Save Current Form */
          <form onSubmit={handleSavePreset} className="space-y-4">
            <div className="bg-indigo-50/70 p-4 rounded-2xl border border-indigo-200 space-y-2">
              <div className="flex items-center gap-2 text-indigo-900 font-black text-xs">
                <Sparkles size={16} />
                <span>လက်ရှိ အစီရင်ခံစာ ချိန်ညှိချက်များအား မှတ်သားသိမ်းဆည်းမည်</span>
              </div>
              <ul className="text-[11px] text-indigo-800 space-y-1 font-medium font-pyidaungsu list-disc list-inside">
                <li>ခေါင်းစဉ်: {currentSettings.tableTitle}</li>
                <li>လက်မှတ်ထိုးအရာရှိ: {currentSettings.officerTitle} {currentSettings.officerName}</li>
                <li>ကော်လံအရေအတွက်: {currentSettings.visibleColumns.length} ခု</li>
                <li>ဒေတာစုစည်းမှု: {currentSettings.dataPool}</li>
                {currentSettings.filterAddressType && currentSettings.filterAddressType !== 'ALL' && (
                  <li>လိပ်စာအမျိုးအစား: {currentSettings.filterAddressType === 'COMPANY' ? 'ကုမ္ပဏီ (Co.Ltd)' : 'အခြားလိပ်စာ (Other)'}</li>
                )}
                {currentSettings.filterAddresses && currentSettings.filterAddresses.length > 0 && (
                  <li>တည်းခိုလိပ်စာ ({currentSettings.filterAddresses.length} ခု): {currentSettings.filterAddresses.slice(0, 3).join(', ')}{currentSettings.filterAddresses.length > 3 ? '...' : ''}</li>
                )}
                {currentSettings.filterNationalities && currentSettings.filterNationalities.length > 0 && (
                  <li>နိုင်ငံအမည်: {currentSettings.filterNationalities.join(', ')}</li>
                )}
                {currentSettings.filterVisaTypes && currentSettings.filterVisaTypes.length > 0 && (
                  <li>ဗီဇာအမျိုးအစား: {currentSettings.filterVisaTypes.join(', ')}</li>
                )}
                {currentSettings.minElapsedDays && (
                  <li>အနည်းဆုံးနေထိုင်ရက်: {currentSettings.minElapsedDays} ရက်အထက်</li>
                )}
              </ul>
            </div>

            <div className="space-y-1.5">
              <label className="text-xs font-black text-slate-700">
                Template အမည် (Preset Name) *
              </label>
              <input
                type="text"
                value={newPresetName}
                onChange={(e) => setNewPresetName(e.target.value)}
                placeholder="ဥပမာ - လစဉ် ဟိုတယ်တည်းခိုသူများ စာရင်း"
                className="w-full px-3.5 py-2.5 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 font-pyidaungsu outline-none font-bold"
                required
              />
            </div>

            <div className="space-y-1.5">
              <label className="text-xs font-black text-slate-700">
                ရှင်းလင်းချက် (Optional Description)
              </label>
              <input
                type="text"
                value={newPresetDesc}
                onChange={(e) => setNewPresetDesc(e.target.value)}
                placeholder="ဥပမာ - နိုင်ငံခြားသား ဧည့်စာရင်းစစ်ဆေးမှုအတွက် အသုံးပြုရန်"
                className="w-full px-3.5 py-2.5 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 font-pyidaungsu outline-none"
              />
            </div>

            <div className="flex justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setActiveTab('list')}
                className="px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition-all cursor-pointer"
              >
                မသိမ်းတော့ပါ
              </button>
              <button
                type="submit"
                className="px-6 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-black rounded-xl shadow-md transition-all flex items-center gap-1.5 cursor-pointer"
              >
                <Check size={14} />
                <span>သိမ်းဆည်းမည် (Save Preset)</span>
              </button>
            </div>
          </form>
        )}
      </div>
    </div>
  );
};
