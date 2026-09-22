import React from 'react';
import { Calendar, ArrowRight, Clock, Sparkles } from 'lucide-react';
import { autoSeparateDob } from './DobNumpadInput';

export const parseAnyDateToObj = (str: string): Date | null => {
  if (!str) return null;
  const trimmed = str.trim();
  if (!trimmed) return null;

  // 1. DD-MM-YYYY or DD/MM/YYYY or DD.MM.YYYY
  const dmyMatch = trimmed.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (dmyMatch) {
    const [, d, m, y] = dmyMatch;
    const dt = new Date(parseInt(y, 10), parseInt(m, 10) - 1, parseInt(d, 10));
    if (!isNaN(dt.getTime())) return dt;
  }

  // 2. YYYY-MM-DD or YYYY/MM/DD or YYYY.MM.DD
  const ymdMatch = trimmed.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:T.*)?$/);
  if (ymdMatch) {
    const [, y, m, d] = ymdMatch;
    const dt = new Date(parseInt(y, 10), parseInt(m, 10) - 1, parseInt(d, 10));
    if (!isNaN(dt.getTime())) return dt;
  }

  const dt = new Date(trimmed);
  if (!isNaN(dt.getTime())) return dt;
  return null;
};

export const formatDateToDDMMYYYYStr = (d: Date): string => {
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  return `${day}-${month}-${year}`;
};

interface StayPeriodInputProps {
  stayFrom: string;
  stayTo: string;
  totalDays: string;
  onChange: (updates: { stayFrom?: string; stayTo?: string; totalDays?: string }) => void;
}

export const StayPeriodInput: React.FC<StayPeriodInputProps> = ({
  stayFrom,
  stayTo,
  totalDays,
  onChange
}) => {
  // Handle Stay From input change
  const handleFromChange = (rawVal: string) => {
    const formatted = autoSeparateDob(rawVal);
    let newTo = stayTo;

    // If formatted is a full date (DD-MM-YYYY) and totalDays exists, auto-calculate stayTo
    if (formatted.length === 10 && totalDays) {
      const numDays = parseInt(totalDays.replace(/\D/g, ''), 10);
      if (!isNaN(numDays) && numDays > 0) {
        const dt = parseAnyDateToObj(formatted);
        if (dt) {
          dt.setDate(dt.getDate() + (numDays - 1));
          newTo = formatDateToDDMMYYYYStr(dt);
        }
      }
    }

    onChange({ stayFrom: formatted, stayTo: newTo });
  };

  // Handle Total Days change
  const handleDaysChange = (rawVal: string) => {
    const digits = rawVal.replace(/\D/g, '');
    let newTo = stayTo;

    if (stayFrom && stayFrom.length === 10 && digits) {
      const numDays = parseInt(digits, 10);
      if (!isNaN(numDays) && numDays > 0) {
        const dt = parseAnyDateToObj(stayFrom);
        if (dt) {
          dt.setDate(dt.getDate() + (numDays - 1));
          newTo = formatDateToDDMMYYYYStr(dt);
        }
      }
    }

    onChange({ totalDays: digits ? `${digits} D` : '', stayTo: newTo });
  };

  // Handle Stay To input change
  const handleToChange = (rawVal: string) => {
    const formatted = autoSeparateDob(rawVal);
    let newDays = totalDays;

    if (stayFrom && stayFrom.length === 10 && formatted.length === 10) {
      const d1 = parseAnyDateToObj(stayFrom);
      const d2 = parseAnyDateToObj(formatted);
      if (d1 && d2) {
        const diffTime = d2.getTime() - d1.getTime();
        const diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24)) + 1;
        newDays = diffDays > 0 ? `${diffDays} D` : '';
      }
    }

    onChange({ stayTo: formatted, totalDays: newDays });
  };

  // Quick preset days click
  const applyPresetDays = (daysCount: number) => {
    const daysStr = `${daysCount} D`;
    let newTo = stayTo;

    // If stayFrom is not set, default to today
    let currentFrom = stayFrom;
    if (!currentFrom || currentFrom.length < 10) {
      currentFrom = formatDateToDDMMYYYYStr(new Date());
    }

    const dt = parseAnyDateToObj(currentFrom);
    if (dt) {
      dt.setDate(dt.getDate() + (daysCount - 1));
      newTo = formatDateToDDMMYYYYStr(dt);
    }

    onChange({ stayFrom: currentFrom, totalDays: daysStr, stayTo: newTo });
  };

  // Quick set Today as Stay From
  const setTodayAsFrom = () => {
    const todayStr = formatDateToDDMMYYYYStr(new Date());
    handleFromChange(todayStr);
  };

  const handleKeyDown = (
    e: React.KeyboardEvent<HTMLInputElement>,
    currentVal: string,
    onValueChange: (val: string) => void
  ) => {
    if (e.key === 'Backspace') {
      const input = e.currentTarget;
      const val = input.value;
      const selStart = input.selectionStart;
      const selEnd = input.selectionEnd;

      if (selStart === selEnd && selStart !== null && selStart > 0 && val[selStart - 1] === '-') {
        e.preventDefault();
        const before = val.slice(0, selStart - 2);
        const after = val.slice(selStart);
        const rawDigits = (before + after).replace(/\D/g, '');
        onValueChange(rawDigits);
      }
    }
  };

  return (
    <div className="space-y-1.5">
      <div className="flex items-center justify-between">
        <label className="input-label text-slate-800 font-bold mb-0 flex items-center gap-1.5 text-xs">
          <Clock size={14} className="text-indigo-600" />
          <span>Stay Period (နေထိုင်ခွင့်ကာလ)</span>
        </label>
        <button
          type="button"
          onClick={setTodayAsFrom}
          className="text-[10px] font-bold px-2 py-0.5 rounded-lg bg-indigo-50 text-indigo-700 border border-indigo-200 hover:bg-indigo-100 transition-colors flex items-center gap-1 cursor-pointer"
          title="စတင်ရက်ကို ယနေ့ရက်စွဲ သတ်မှတ်မည်"
        >
          <Sparkles size={11} className="text-indigo-500" />
          <span>ယနေ့ရက်စွဲ (Today)</span>
        </button>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-12 gap-2 items-center bg-slate-50/70 p-2 rounded-xl border border-slate-200">
        {/* Stay From */}
        <div className="sm:col-span-5">
          <div className="text-[10px] font-bold text-slate-500 mb-0.5 flex items-center gap-1">
            <span>စတင်ရက် (From)</span>
          </div>
          <div className="relative">
            <input
              type="text"
              inputMode="numeric"
              pattern="[0-9\-]*"
              value={stayFrom || ''}
              onChange={(e) => handleFromChange(e.target.value)}
              onKeyDown={(e) => handleKeyDown(e, stayFrom, handleFromChange)}
              onBlur={() => {
                if (stayFrom) handleFromChange(autoSeparateDob(stayFrom));
              }}
              className="input-field text-gray-900 font-bold font-mono tracking-wider text-xs bg-white border-indigo-200 focus:border-indigo-600 h-9"
              placeholder="DD-MM-YYYY"
              maxLength={10}
            />
          </div>
        </div>

        {/* Total Days */}
        <div className="sm:col-span-2 text-center">
          <div className="text-[10px] font-bold text-indigo-700 mb-0.5">
            <span>ရက်ပေါင်း</span>
          </div>
          <input
            type="text"
            inputMode="numeric"
            value={totalDays ? totalDays.replace(/[^0-9]/g, '') : ''}
            onChange={(e) => handleDaysChange(e.target.value)}
            className="input-field text-center font-mono font-black text-indigo-900 text-xs bg-indigo-50/70 border-indigo-300 focus:border-indigo-600 h-9 w-full"
            placeholder="Days"
            maxLength={4}
          />
        </div>

        {/* Stay To */}
        <div className="sm:col-span-5">
          <div className="text-[10px] font-bold text-slate-500 mb-0.5 flex items-center gap-1">
            <ArrowRight size={11} className="text-slate-400" />
            <span>ကုန်ဆုံးရက် (To)</span>
          </div>
          <div className="relative">
            <input
              type="text"
              inputMode="numeric"
              pattern="[0-9\-]*"
              value={stayTo || ''}
              onChange={(e) => handleToChange(e.target.value)}
              onKeyDown={(e) => handleKeyDown(e, stayTo, handleToChange)}
              onBlur={() => {
                if (stayTo) handleToChange(autoSeparateDob(stayTo));
              }}
              className="input-field text-gray-900 font-bold font-mono tracking-wider text-xs bg-white border-indigo-200 focus:border-indigo-600 h-9"
              placeholder="DD-MM-YYYY"
              maxLength={10}
            />
          </div>
        </div>
      </div>

      {/* Quick Preset Days Pills */}
      <div className="flex items-center gap-1.5 flex-wrap pt-0.5">
        <span className="text-[10px] font-semibold text-slate-500">ရက်အမြန်ရွေး:</span>
        {[14, 30, 70, 90, 180, 365].map((d) => {
          const isSelected = totalDays && parseInt(totalDays.replace(/\D/g, ''), 10) === d;
          return (
            <button
              key={d}
              type="button"
              onClick={() => applyPresetDays(d)}
              className={`text-[10px] font-bold px-2 py-0.5 rounded-md border transition-all cursor-pointer ${
                isSelected
                  ? 'bg-indigo-600 text-white border-indigo-600 shadow-xs'
                  : 'bg-white text-slate-700 border-slate-200 hover:bg-indigo-50 hover:text-indigo-700 hover:border-indigo-300'
              }`}
            >
              {d === 365 ? '1 Year (365D)' : `${d}D`}
            </button>
          );
        })}
      </div>
    </div>
  );
};
