import React, { useState, useRef, useEffect } from 'react';
import { Calendar, Delete, X, Check, Calculator } from 'lucide-react';

/**
 * Universal DOB Formatter & Auto-Separator
 * Converts any date format (YYYY-MM-DD, YYYY/MM/DD, DD/MM/YYYY, raw digits) into DD-MM-YYYY
 */
export const autoSeparateDob = (raw: string): string => {
  if (!raw) return '';
  const trimmed = raw.trim();
  if (!trimmed) return '';

  // 1. Detect standard ISO/Autofill date format: YYYY-MM-DD, YYYY/MM/DD, YYYY.MM.DD
  const ymdMatch = trimmed.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:T.*)?$/);
  if (ymdMatch) {
    const [, y, m, d] = ymdMatch;
    return `${d.padStart(2, '0')}-${m.padStart(2, '0')}-${y}`;
  }

  // 2. Detect formatted DD-MM-YYYY, DD/MM/YYYY, DD.MM.YYYY
  const dmyMatch = trimmed.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (dmyMatch) {
    const [, d, m, y] = dmyMatch;
    return `${d.padStart(2, '0')}-${m.padStart(2, '0')}-${y}`;
  }

  // 3. Handle pure digit typing / input without formatting
  const digits = raw.replace(/\D/g, '').slice(0, 8);
  if (!digits) return '';

  // If 8 digits pasted/autofilled starting with year 19xx or 20xx (e.g. 19951225 -> 25-12-1995)
  if (digits.length === 8) {
    const first4 = parseInt(digits.slice(0, 4), 10);
    if (first4 >= 1900 && first4 <= 2099) {
      const year = digits.slice(0, 4);
      const month = digits.slice(4, 6);
      const day = digits.slice(6, 8);
      return `${day}-${month}-${year}`;
    }
  }

  // Progressive typing for DD-MM-YYYY
  if (digits.length <= 2) {
    return digits;
  } else if (digits.length <= 4) {
    return `${digits.slice(0, 2)}-${digits.slice(2)}`;
  } else {
    return `${digits.slice(0, 2)}-${digits.slice(2, 4)}-${digits.slice(4, 8)}`;
  }
};

interface DobNumpadInputProps {
  value: string;
  onChange: (formattedValue: string) => void;
  label?: string;
  placeholder?: string;
  required?: boolean;
}

export const DobNumpadInput: React.FC<DobNumpadInputProps> = ({
  value,
  onChange,
  label = 'မွေးသက္ကရာဇ် (Date of Birth / DOB)',
  placeholder = 'DD-MM-YYYY (e.g. 25-12-1995)',
  required = false
}) => {
  const [showNumpad, setShowNumpad] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  // Close numpad when clicking outside
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setShowNumpad(false);
      }
    };
    if (showNumpad) {
      document.addEventListener('mousedown', handleClickOutside);
    }
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [showNumpad]);

  const handleDigitPress = (digit: string) => {
    const rawDigits = (value || '').replace(/\D/g, '');
    if (rawDigits.length >= 8) return;
    const newDigits = rawDigits + digit;
    onChange(autoSeparateDob(newDigits));
  };

  const handleBackspace = () => {
    const rawDigits = (value || '').replace(/\D/g, '');
    if (rawDigits.length === 0) return;
    const newDigits = rawDigits.slice(0, -1);
    onChange(autoSeparateDob(newDigits));
  };

  const handleClear = () => {
    onChange('');
  };

  const handleDirectInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const inputVal = e.target.value;
    if (!inputVal) {
      onChange('');
      return;
    }
    onChange(autoSeparateDob(inputVal));
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Backspace') {
      const input = e.currentTarget;
      const val = input.value;
      const selStart = input.selectionStart;
      const selEnd = input.selectionEnd;

      // If cursor is immediately after a hyphen, remove the hyphen and preceding digit
      if (selStart === selEnd && selStart !== null && selStart > 0 && val[selStart - 1] === '-') {
        e.preventDefault();
        const before = val.slice(0, selStart - 2);
        const after = val.slice(selStart);
        const rawDigits = (before + after).replace(/\D/g, '');
        onChange(autoSeparateDob(rawDigits));
      }
    }
  };

  const handleBlur = () => {
    if (value) {
      onChange(autoSeparateDob(value));
    }
  };

  return (
    <div className="relative" ref={containerRef}>
      <div className="flex items-center justify-between mb-1">
        <label className="input-label text-slate-700 font-bold mb-0 flex items-center gap-1.5 text-xs">
          <Calendar size={14} className="text-indigo-600" />
          <span>{label}</span>
          {required && <span className="text-red-500">*</span>}
        </label>
        <button
          type="button"
          onClick={() => setShowNumpad(!showNumpad)}
          className={`text-[10px] font-bold px-2 py-0.5 rounded-lg border transition-all flex items-center gap-1 cursor-pointer ${
            showNumpad
              ? 'bg-indigo-900 text-white border-indigo-900 shadow-xs'
              : 'bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100 hover:text-indigo-700'
          }`}
          title="On-Screen Touch Pad ဖွင့်/ပိတ်ရန်"
        >
          <Calculator size={12} />
          <span>{showNumpad ? 'Pad ပိတ်မည်' : '🔢 On-Screen Pad'}</span>
        </button>
      </div>

      <div className="relative">
        <input
          type="text"
          inputMode="numeric"
          pattern="[0-9\-]*"
          value={value || ''}
          onChange={handleDirectInputChange}
          onKeyDown={handleKeyDown}
          onBlur={handleBlur}
          className="input-field text-gray-900 font-bold font-mono tracking-wider text-xs pr-10 bg-white border-indigo-200 focus:border-indigo-600 focus:ring-2 focus:ring-indigo-100 h-10"
          placeholder={placeholder}
          maxLength={10}
        />
        {value && (
          <button
            type="button"
            onClick={handleClear}
            className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-red-500 p-1 cursor-pointer transition-colors"
            title="Clear DOB"
          >
            <X size={15} />
          </button>
        )}
      </div>

      {/* Optional On-Screen Numpad Popover (Only when explicitly toggled via button) */}
      {showNumpad && (
        <div className="absolute z-50 left-0 right-0 sm:left-auto sm:right-0 mt-2 p-3 bg-slate-900 text-white rounded-2xl shadow-2xl border border-slate-700 w-full sm:w-72 animate-in fade-in-50 zoom-in-95 duration-150">
          <div className="flex items-center justify-between pb-2 mb-2 border-b border-slate-700/80">
            <div className="flex items-center gap-1.5">
              <span className="text-xs font-bold text-indigo-300">🔢 DOB On-Screen Numpad</span>
            </div>
            <div className="text-xs font-mono font-black text-amber-300 bg-slate-800 px-2 py-0.5 rounded border border-slate-700">
              {value || 'DD-MM-YYYY'}
            </div>
          </div>

          <div className="grid grid-cols-3 gap-1.5 mb-2">
            {['1', '2', '3', '4', '5', '6', '7', '8', '9'].map((num) => (
              <button
                key={num}
                type="button"
                onClick={() => handleDigitPress(num)}
                className="h-10 bg-slate-800 hover:bg-indigo-600 active:scale-95 text-white font-mono font-black text-base rounded-xl border border-slate-700 shadow-xs transition-all flex items-center justify-center cursor-pointer"
              >
                {num}
              </button>
            ))}
            <button
              type="button"
              onClick={handleClear}
              className="h-10 bg-rose-900/60 hover:bg-rose-700 active:scale-95 text-rose-200 font-bold text-xs rounded-xl border border-rose-800/80 transition-all flex items-center justify-center cursor-pointer"
              title="Clear all"
            >
              Clear
            </button>
            <button
              type="button"
              onClick={() => handleDigitPress('0')}
              className="h-10 bg-slate-800 hover:bg-indigo-600 active:scale-95 text-white font-mono font-black text-base rounded-xl border border-slate-700 shadow-xs transition-all flex items-center justify-center cursor-pointer"
            >
              0
            </button>
            <button
              type="button"
              onClick={handleBackspace}
              className="h-10 bg-amber-900/60 hover:bg-amber-700 active:scale-95 text-amber-200 font-bold text-xs rounded-xl border border-amber-800/80 transition-all flex items-center justify-center cursor-pointer"
              title="Backspace"
            >
              <Delete size={16} />
            </button>
          </div>

          <button
            type="button"
            onClick={() => setShowNumpad(false)}
            className="w-full py-1.5 bg-indigo-600 hover:bg-indigo-500 active:scale-98 text-white font-bold text-xs rounded-xl transition-all flex items-center justify-center gap-1.5 shadow-xs cursor-pointer"
          >
            <Check size={14} />
            <span>ပြီးစီး (Done)</span>
          </button>
        </div>
      )}
    </div>
  );
};
