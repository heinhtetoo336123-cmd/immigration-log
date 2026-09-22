import React, { useState, useMemo, useRef, useEffect } from 'react';
import { Building2, Home, MapPin, Search, Check, Sparkles, ChevronDown, X } from 'lucide-react';
import { MasterItem } from '../types';
import { isCompanyAddress } from '../utils/addressUtils';

interface MasterLocationInputProps {
  locationValue: string;
  onLocationChange: (val: string) => void;
  descriptionValue: string;
  onDescriptionChange: (val: string) => void;
  masterData?: MasterItem[];
  locationLabel?: string;
  descriptionLabel?: string;
  locationPlaceholder?: string;
  descriptionPlaceholder?: string;
  required?: boolean;
  disabled?: boolean;
  idPrefix?: string;
  showQuickPills?: boolean;
}

export const MasterLocationInput: React.FC<MasterLocationInputProps> = ({
  locationValue,
  onLocationChange,
  descriptionValue,
  onDescriptionChange,
  masterData = [],
  locationLabel = 'တည်းခိုရာ နေရာ / ကုမ္ပဏီ (Stay Location / Company)',
  descriptionLabel = 'လိပ်စာ အသေးစိတ် (Stay Description / Detail)',
  locationPlaceholder = 'ဥပမာ- Tidy Co., Ltd (သို့) ဟိုတယ်အမည် ရိုက်ထည့်ပါ...',
  descriptionPlaceholder = 'အဆောင်၊ အခန်းအမှတ်၊ ရပ်ကွက်၊ မြို့နယ်...',
  required = false,
  disabled = false,
  idPrefix = 'loc_input',
  showQuickPills = true
}) => {
  const [isOpen, setIsOpen] = useState(false);
  const [filterCategory, setFilterCategory] = useState<'ALL' | 'COMPANY' | 'OTHER'>('ALL');
  const [justAutofilled, setJustAutofilled] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  // Extract all unique Stay items from masterData
  const stayMasterItems = useMemo(() => {
    const seen = new Set<string>();
    const list: Array<{ name: string; linkedValue: string; isCo: boolean }> = [];

    masterData
      .filter(m => m && m.type === 'Stay' && m.name && m.name.trim())
      .forEach(m => {
        const cleanName = m.name.trim();
        const key = cleanName.toLowerCase();
        if (!seen.has(key)) {
          seen.add(key);
          const isCo = isCompanyAddress(cleanName, masterData);
          list.push({
            name: cleanName,
            linkedValue: m.linkedValue ? m.linkedValue.trim() : '',
            isCo
          });
        }
      });

    return list.sort((a, b) => a.name.localeCompare(b.name));
  }, [masterData]);

  // Filter master items according to search query and category
  const filteredSuggestions = useMemo(() => {
    const q = (locationValue || '').toLowerCase().trim();
    return stayMasterItems.filter(item => {
      if (filterCategory === 'COMPANY' && !item.isCo) return false;
      if (filterCategory === 'OTHER' && item.isCo) return false;

      if (!q) return true;
      const matchName = item.name.toLowerCase().includes(q);
      const matchDesc = item.linkedValue.toLowerCase().includes(q);
      return matchName || matchDesc;
    });
  }, [stayMasterItems, locationValue, filterCategory]);

  // Check if current value matches any master item
  const matchedMasterItem = useMemo(() => {
    if (!locationValue?.trim()) return null;
    const clean = locationValue.trim().toLowerCase();
    return stayMasterItems.find(m => m.name.toLowerCase() === clean) || null;
  }, [stayMasterItems, locationValue]);

  // Close dropdown on outside click
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleSelectMaster = (item: { name: string; linkedValue: string }) => {
    onLocationChange(item.name);
    if (item.linkedValue) {
      onDescriptionChange(item.linkedValue);
    }
    setJustAutofilled(true);
    setIsOpen(false);
    setTimeout(() => setJustAutofilled(false), 3500);
  };

  const handleLocationInputChange = (val: string) => {
    onLocationChange(val);
    setIsOpen(true);

    // If user typed exact match of a master location, auto-fill description if currently blank
    const exact = stayMasterItems.find(m => m.name.toLowerCase() === val.trim().toLowerCase());
    if (exact && exact.linkedValue && !descriptionValue.trim()) {
      onDescriptionChange(exact.linkedValue);
      setJustAutofilled(true);
      setTimeout(() => setJustAutofilled(false), 3000);
    }
  };

  const datalistId = `${idPrefix}_stay_datalist`;

  return (
    <div ref={containerRef} className="space-y-3 relative">
      {/* Native datalist for browser fallback */}
      <datalist id={datalistId}>
        {stayMasterItems.map(m => (
          <option key={m.name} value={m.name}>
            {m.linkedValue ? `${m.name} (${m.linkedValue})` : m.name}
          </option>
        ))}
      </datalist>

      {/* Location / Company Name Input */}
      <div>
        <div className="flex items-center justify-between mb-1.5">
          <label className="block text-[11px] font-black uppercase text-slate-700">
            {locationLabel} {required && <span className="text-rose-500">*</span>}
          </label>
          {stayMasterItems.length > 0 && (
            <button
              type="button"
              onClick={() => setIsOpen(prev => !prev)}
              className="text-[10px] font-bold text-indigo-600 hover:text-indigo-800 flex items-center gap-1 cursor-pointer"
            >
              <Sparkles size={12} className="text-amber-500" />
              <span>Master Data မှ ရွေးချယ်ရန် ({stayMasterItems.length})</span>
              <ChevronDown size={12} className={`transition-transform ${isOpen ? 'rotate-180' : ''}`} />
            </button>
          )}
        </div>

        <div className="relative">
          <div className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none">
            {matchedMasterItem?.isCo ? (
              <Building2 size={16} className="text-purple-600" />
            ) : matchedMasterItem ? (
              <Home size={16} className="text-blue-600" />
            ) : (
              <MapPin size={16} />
            )}
          </div>

          <input
            type="text"
            list={datalistId}
            value={locationValue}
            onChange={(e) => handleLocationInputChange(e.target.value)}
            onFocus={() => setIsOpen(true)}
            placeholder={locationPlaceholder}
            required={required}
            disabled={disabled}
            className="w-full pl-9 pr-10 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:bg-white transition-all"
          />

          {locationValue && (
            <button
              type="button"
              onClick={() => {
                onLocationChange('');
                setIsOpen(true);
              }}
              className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 p-1 cursor-pointer"
            >
              <X size={14} />
            </button>
          )}
        </div>

        {/* Master Match & Autofill Status Pill */}
        {justAutofilled && matchedMasterItem?.linkedValue && (
          <div className="mt-1.5 flex items-center gap-1.5 text-[11px] font-bold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2.5 py-1 rounded-lg animate-fadeIn">
            <Check size={13} className="shrink-0" />
            <span>Master Data မှ လိပ်စာအသေးစိတ်အား အလိုအလျောက် ဖြည့်စွက်ပြီးပါပြီ</span>
          </div>
        )}
      </div>

      {/* Floating Master Data Suggestion / Selection Dropdown */}
      {isOpen && stayMasterItems.length > 0 && (
        <div className="absolute left-0 right-0 z-50 bg-white border border-slate-200 rounded-2xl shadow-2xl overflow-hidden animate-fadeIn mt-1 max-h-72 flex flex-col">
          {/* Header & Filter Tabs */}
          <div className="p-2.5 bg-slate-100 border-b border-slate-200 flex items-center justify-between gap-2 shrink-0">
            <div className="flex items-center gap-1 text-[11px] font-black text-slate-700">
              <Sparkles size={13} className="text-amber-500" />
              <span>တည်းခိုနေရာ Master Data စာရင်း ({filteredSuggestions.length})</span>
            </div>
            <div className="flex items-center gap-1 bg-white p-0.5 rounded-lg border border-slate-200 text-[10px] font-bold">
              <button
                type="button"
                onClick={() => setFilterCategory('ALL')}
                className={`px-2 py-0.5 rounded ${filterCategory === 'ALL' ? 'bg-indigo-600 text-white' : 'text-slate-600'}`}
              >
                အားလုံး
              </button>
              <button
                type="button"
                onClick={() => setFilterCategory('COMPANY')}
                className={`px-2 py-0.5 rounded ${filterCategory === 'COMPANY' ? 'bg-purple-600 text-white' : 'text-slate-600'}`}
              >
                🏢 ကုမ္ပဏီ
              </button>
              <button
                type="button"
                onClick={() => setFilterCategory('OTHER')}
                className={`px-2 py-0.5 rounded ${filterCategory === 'OTHER' ? 'bg-blue-600 text-white' : 'text-slate-600'}`}
              >
                🏨 ဟိုတယ်/အခြား
              </button>
            </div>
          </div>

          {/* List of Suggestions */}
          <div className="overflow-y-auto divide-y divide-slate-100 p-1">
            {filteredSuggestions.length === 0 ? (
              <div className="p-4 text-center text-xs text-slate-400 font-bold">
                ကိုက်ညီသော Master Data မရှိပါ (အသစ်အဖြစ် တိုက်ရိုက်ရိုက်ထည့်နိုင်ပါသည်)
              </div>
            ) : (
              filteredSuggestions.map(item => {
                const isSelected = item.name.toLowerCase() === (locationValue || '').toLowerCase().trim();
                return (
                  <div
                    key={'sugg_' + item.name}
                    onClick={() => handleSelectMaster(item)}
                    className={`p-2.5 rounded-xl text-xs cursor-pointer flex items-center justify-between gap-3 transition-colors ${
                      isSelected 
                        ? 'bg-indigo-50 border border-indigo-200 text-indigo-950 font-black' 
                        : 'hover:bg-slate-50 text-slate-800'
                    }`}
                  >
                    <div className="flex items-center gap-2.5 min-w-0">
                      <div className={`p-1.5 rounded-lg shrink-0 ${
                        item.isCo ? 'bg-purple-100 text-purple-700' : 'bg-blue-100 text-blue-700'
                      }`}>
                        {item.isCo ? <Building2 size={15} /> : <Home size={15} />}
                      </div>
                      <div className="min-w-0">
                        <div className="font-bold truncate text-slate-900 flex items-center gap-1.5">
                          <span>{item.name}</span>
                          <span className={`text-[9px] font-black px-1.5 py-0.2 rounded uppercase ${
                            item.isCo ? 'bg-purple-100 text-purple-800' : 'bg-blue-100 text-blue-800'
                          }`}>
                            {item.isCo ? '🏢 Co.,Ltd' : '🏨 Hotel'}
                          </span>
                        </div>
                        {item.linkedValue && (
                          <div className="text-[11px] text-slate-400 font-medium truncate mt-0.5">
                            🔗 {item.linkedValue}
                          </div>
                        )}
                      </div>
                    </div>

                    {item.linkedValue && (
                      <span className="shrink-0 text-[10px] bg-slate-100 text-slate-600 px-2 py-0.5 rounded-md font-bold">
                        Autofill Available
                      </span>
                    )}
                  </div>
                );
              })
            )}
          </div>
        </div>
      )}

      {/* Quick Selection Pills (Top Master Locations) */}
      {showQuickPills && stayMasterItems.length > 0 && !isOpen && (
        <div className="flex items-center gap-1.5 flex-wrap pt-0.5">
          <span className="text-[10px] font-bold text-slate-400 flex items-center gap-1">
            <Sparkles size={11} className="text-amber-500" />
            <span>အသုံးများ:</span>
          </span>
          {stayMasterItems.slice(0, 6).map(m => (
            <button
              key={'pill_' + m.name}
              type="button"
              onClick={() => handleSelectMaster(m)}
              className={`text-[10px] font-bold px-2 py-0.5 rounded-lg border transition-all cursor-pointer ${
                m.name.toLowerCase() === (locationValue || '').toLowerCase().trim()
                  ? 'bg-indigo-600 text-white border-indigo-600 shadow-xs'
                  : 'bg-white text-slate-700 border-slate-200 hover:border-indigo-300 hover:bg-slate-50'
              }`}
            >
              {m.isCo ? '🏢 ' : '🏨 '}{m.name}
            </button>
          ))}
        </div>
      )}

      {/* Stay Description / Detail Input */}
      <div>
        <label className="block text-[11px] font-black uppercase text-slate-700 mb-1.5">
          {descriptionLabel}
        </label>
        <input
          type="text"
          value={descriptionValue}
          onChange={(e) => onDescriptionChange(e.target.value)}
          placeholder={descriptionPlaceholder}
          disabled={disabled}
          className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:bg-white transition-all"
        />
      </div>
    </div>
  );
};
