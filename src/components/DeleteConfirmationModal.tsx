import React, { useState, useEffect } from 'react';
import { ShieldAlert, Trash2, AlertTriangle, X, CheckCircle2 } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

interface DeleteConfirmationModalProps {
  isOpen: boolean;
  onClose: () => void;
  onConfirm: (reason: string) => void;
  title: string;
  itemDetails?: {
    primaryText: string;
    secondaryText?: string;
    infoList?: Array<{ label: string; value: string }>;
  };
  warningText?: string;
  isBulkAction?: boolean;
}

export const DeleteConfirmationModal: React.FC<DeleteConfirmationModalProps> = ({
  isOpen,
  onClose,
  onConfirm,
  title,
  itemDetails,
  warningText,
  isBulkAction = false,
}) => {
  const [inputText, setInputText] = useState('');
  const [reason, setReason] = useState('');

  useEffect(() => {
    if (isOpen) {
      setInputText('');
      setReason('');
    }
  }, [isOpen]);

  if (!isOpen) return null;

  const isConfirmed = inputText.trim().toUpperCase() === 'CONFIRMED';

  const handleExecuteDelete = (e: React.FormEvent) => {
    e.preventDefault();
    if (!isConfirmed) return;
    onConfirm(reason.trim() || 'Manual Delete with Security Safeguard');
    onClose();
  };

  return (
    <AnimatePresence>
      <div className="fixed inset-0 z-[300] flex items-center justify-center p-4 select-text no-print">
        {/* Backdrop */}
        <motion.div
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          onClick={onClose}
          className="absolute inset-0 bg-slate-950/80 backdrop-blur-md"
        />

        {/* Modal Dialog */}
        <motion.div
          initial={{ scale: 0.9, opacity: 0, y: 20 }}
          animate={{ scale: 1, opacity: 1, y: 0 }}
          exit={{ scale: 0.9, opacity: 0, y: 20 }}
          className="relative bg-white dark:bg-slate-900 w-full max-w-lg rounded-3xl overflow-hidden shadow-2xl border-2 border-rose-500/40 text-slate-900 dark:text-white"
        >
          {/* Header */}
          <div className="bg-rose-950 p-6 border-b border-rose-800/60 flex items-start justify-between gap-4 text-white">
            <div className="flex items-center gap-3">
              <div className="w-12 h-12 rounded-2xl bg-rose-900/80 text-rose-200 flex items-center justify-center shrink-0 border border-rose-700/50 shadow-inner animate-pulse">
                <ShieldAlert size={26} />
              </div>
              <div>
                <span className="text-[10px] font-black uppercase tracking-widest text-rose-300 block">
                  Security Deletion Safeguard
                </span>
                <h3 className="text-lg font-black text-white leading-tight mt-0.5">
                  {title}
                </h3>
              </div>
            </div>
            <button
              onClick={onClose}
              className="p-1.5 rounded-full hover:bg-rose-900/60 text-rose-300 hover:text-white transition-colors cursor-pointer"
            >
              <X size={18} />
            </button>
          </div>

          <form onSubmit={handleExecuteDelete} className="p-6 space-y-5">
            {/* Warning Banner */}
            <div className="bg-rose-50 dark:bg-rose-950/40 border border-rose-200 dark:border-rose-800/80 p-4 rounded-2xl flex items-start gap-3">
              <AlertTriangle size={20} className="text-rose-600 dark:text-rose-400 shrink-0 mt-0.5" />
              <div className="text-xs font-semibold text-rose-950 dark:text-rose-200 leading-relaxed">
                {warningText || (
                  <span>
                    ယခုပယ်ဖျက်လိုက်သည့် မှတ်တမ်းသည် အမှိုက်ပုံး (Recycle Bin) သို့ အလိုအလျောက် ရောက်ရှိသွားမည်ဖြစ်သည်။ လုံးဝ မှားယွင်းမှုမရှိစေရန်အတွက် အောက်ပါ အကွက်တွင် <strong className="font-black underline text-rose-700 dark:text-rose-300">"CONFIRMED"</strong> ဟု ရိုက်ထည့်ပေးပါ။
                  </span>
                )}
              </div>
            </div>

            {/* Target Item Details Box */}
            {itemDetails && (
              <div className="bg-slate-50 dark:bg-slate-800/60 border border-slate-200 dark:border-slate-700/80 p-4 rounded-2xl space-y-2 text-xs">
                <div className="font-mono text-[10px] font-black uppercase text-slate-400 tracking-wider">
                  Target Deletion Info:
                </div>
                <div className="text-sm font-black text-slate-900 dark:text-white uppercase tracking-tight">
                  {itemDetails.primaryText}
                </div>
                {itemDetails.secondaryText && (
                  <div className="text-xs font-bold text-slate-600 dark:text-slate-300 font-mono">
                    {itemDetails.secondaryText}
                  </div>
                )}
                {itemDetails.infoList && itemDetails.infoList.length > 0 && (
                  <div className="grid grid-cols-2 gap-2 pt-2 border-t border-slate-200 dark:border-slate-700/60 text-[11px]">
                    {itemDetails.infoList.map((info, idx) => (
                      <div key={idx}>
                        <span className="text-slate-400 font-medium block text-[10px] uppercase">{info.label}</span>
                        <span className="font-bold text-slate-800 dark:text-slate-200">{info.value}</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}

            {/* Optional Reason Field */}
            <div>
              <label className="block text-[11px] font-bold text-slate-500 dark:text-slate-400 uppercase tracking-wider mb-1">
                Deletion Reason / မှတ်ချက် (Optional)
              </label>
              <input
                type="text"
                value={reason}
                onChange={(e) => setReason(e.target.value)}
                placeholder="e.g. Duplicate entry, Wrong passport, Officer correction..."
                className="w-full px-3.5 py-2 rounded-xl bg-slate-50 dark:bg-slate-800 border border-slate-300 dark:border-slate-700 text-xs font-semibold focus:outline-none focus:ring-2 focus:ring-rose-500"
              />
            </div>

            {/* CONFIRMED Input Field */}
            <div className="space-y-1.5">
              <label className="block text-xs font-black uppercase tracking-wider text-rose-700 dark:text-rose-400">
                To confirm, type <span className="underline decoration-2 font-mono">CONFIRMED</span> below:
              </label>
              <div className="relative">
                <input
                  type="text"
                  value={inputText}
                  onChange={(e) => setInputText(e.target.value)}
                  placeholder="CONFIRMED"
                  autoFocus
                  className="w-full px-4 py-3 rounded-2xl bg-white dark:bg-slate-800 border-2 border-rose-300 dark:border-rose-700 text-center font-mono font-black text-base uppercase tracking-widest text-slate-900 dark:text-white focus:outline-none focus:ring-4 focus:ring-rose-500/20"
                />
                {isConfirmed && (
                  <CheckCircle2 size={20} className="absolute right-3.5 top-3.5 text-emerald-500 animate-in zoom-in-50" />
                )}
              </div>
              <p className="text-[10px] text-slate-400 text-center font-medium">
                (CONFIRMED ဟု စာလုံးအကြီးဖြင့် တိကျစွာ ရိုက်ထည့်ပါ)
              </p>
            </div>

            {/* Action Buttons */}
            <div className="flex items-center gap-3 pt-2">
              <button
                type="button"
                onClick={onClose}
                className="flex-1 py-3 px-4 rounded-xl bg-slate-100 hover:bg-slate-200 dark:bg-slate-800 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-300 text-xs font-black uppercase tracking-wider transition-all cursor-pointer"
              >
                မလုပ်ဆောင်ပါ (Cancel)
              </button>
              <button
                type="submit"
                disabled={!isConfirmed}
                className={`flex-1 py-3 px-4 rounded-xl font-black text-xs uppercase tracking-wider transition-all flex items-center justify-center gap-2 shadow-lg cursor-pointer ${
                  isConfirmed
                    ? 'bg-rose-600 hover:bg-rose-700 text-white shadow-rose-900/30 scale-102 active:scale-95'
                    : 'bg-slate-200 dark:bg-slate-800 text-slate-400 dark:text-slate-600 cursor-not-allowed shadow-none border border-slate-300 dark:border-slate-700'
                }`}
              >
                <Trash2 size={16} />
                <span>Move to Trash</span>
              </button>
            </div>
          </form>
        </motion.div>
      </div>
    </AnimatePresence>
  );
};
