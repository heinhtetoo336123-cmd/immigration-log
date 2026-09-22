import React from 'react';
import { Columns, CheckSquare, Square, RotateCcw, X, Eye, EyeOff } from 'lucide-react';
import { CRTColumnId } from '../types';
import { CRT_COLUMNS_DEFINITION, ALL_COLUMN_IDS } from '../utils/crtPresets';

interface CustomReportColumnModalProps {
  isOpen: boolean;
  onClose: () => void;
  visibleColumns: CRTColumnId[];
  onToggleColumn: (colId: CRTColumnId) => void;
  onSelectAll: () => void;
  onResetDefault: () => void;
}

export const CustomReportColumnModal: React.FC<CustomReportColumnModalProps> = ({
  isOpen,
  onClose,
  visibleColumns,
  onToggleColumn,
  onSelectAll,
  onResetDefault
}) => {
  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[600] flex items-center justify-center p-4 bg-slate-900/80 backdrop-blur-xs animate-in fade-in duration-150">
      <div className="bg-white rounded-3xl p-6 max-w-lg w-full shadow-2xl border border-slate-200 space-y-5">
        {/* Header */}
        <div className="flex items-center justify-between border-b border-slate-100 pb-3">
          <div className="flex items-center gap-2.5">
            <div className="p-2 bg-indigo-100 text-indigo-700 rounded-xl">
              <Columns size={20} />
            </div>
            <div>
              <h3 className="text-base font-black text-slate-900 uppercase">
                Column Visibility Settings
              </h3>
              <p className="text-xs text-slate-500 font-medium">
                ဇယားနှင့် ပုံနှိပ်စာရွက်တွင် ပြသမည့် ကော်လံများ စိတ်ကြိုက် ဖွင့်/ပိတ်ပါ
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

        {/* Quick Actions */}
        <div className="flex items-center justify-between text-xs font-bold text-slate-600 bg-slate-50 p-2.5 rounded-xl border border-slate-200">
          <span>ရွေးချယ်ထားသော ကော်လံ ({visibleColumns.length}/{CRT_COLUMNS_DEFINITION.length}) ခု</span>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onSelectAll}
              className="text-indigo-600 hover:text-indigo-800 hover:underline cursor-pointer"
            >
              အားလုံးဖွင့်မည် (All)
            </button>
            <span>•</span>
            <button
              type="button"
              onClick={onResetDefault}
              className="text-slate-600 hover:text-slate-800 flex items-center gap-1 hover:underline cursor-pointer"
            >
              <RotateCcw size={11} /> မူလအတိုင်း (Default)
            </button>
          </div>
        </div>

        {/* Columns Grid */}
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5 max-h-[340px] overflow-y-auto pr-1">
          {CRT_COLUMNS_DEFINITION.map((col, idx) => {
            const isVisible = visibleColumns.includes(col.id);
            return (
              <label
                key={col.id}
                onClick={() => onToggleColumn(col.id)}
                className={`flex items-center justify-between p-3 rounded-2xl border transition-all cursor-pointer select-none ${
                  isVisible
                    ? 'bg-indigo-50/60 border-indigo-300 text-indigo-950 font-bold shadow-xs'
                    : 'bg-slate-50 border-slate-200 text-slate-400 opacity-70'
                }`}
              >
                <div className="flex items-center gap-2.5 truncate">
                  <span className="w-5 h-5 rounded-md bg-white border border-slate-300 flex items-center justify-center text-[10px] font-mono text-slate-500 shrink-0">
                    {idx + 1}
                  </span>
                  <span className="text-xs truncate font-pyidaungsu">
                    {col.label}
                  </span>
                </div>
                <div className="shrink-0 text-indigo-600 pl-2">
                  {isVisible ? <CheckSquare size={18} /> : <Square size={18} className="text-slate-400" />}
                </div>
              </label>
            );
          })}
        </div>

        {/* Footer */}
        <div className="flex justify-end pt-2 border-t border-slate-100">
          <button
            type="button"
            onClick={onClose}
            className="px-6 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-black rounded-xl shadow-md transition-all cursor-pointer"
          >
            ပြီးပါပြီ (Apply & Close)
          </button>
        </div>
      </div>
    </div>
  );
};
