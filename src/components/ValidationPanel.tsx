import React from 'react';
import { AlertTriangle, CheckCircle2 } from 'lucide-react';
import type { MissingFieldItem } from '../types/memo';

interface ValidationPanelProps {
  missingFields: MissingFieldItem[];
  overrides: Record<string, string>;
  onOverrideChange: (code: string, value: string) => void;
}

export const ValidationPanel: React.FC<ValidationPanelProps> = ({
  missingFields,
  overrides,
  onOverrideChange,
}) => {
  if (missingFields.length === 0) {
    return (
      <div className="bg-white border border-emerald-200 rounded-lg p-4 flex items-center gap-3">
        <CheckCircle2 className="w-5 h-5 text-emerald-700 shrink-0" />
        <div>
          <p className="text-sm font-semibold text-slate-900">
            All IC Memo placeholders resolved
          </p>
          <p className="text-xs text-slate-600">
            Every required field is populated and ready for .docx export.
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="bg-white border border-amber-300 rounded-lg overflow-hidden">
      <div className="px-4 py-3 bg-amber-50 border-b border-amber-200 flex items-center justify-between gap-4">
        <div className="flex items-center gap-2.5">
          <AlertTriangle className="w-4 h-4 text-amber-700 shrink-0" />
          <h3 className="text-sm font-semibold text-slate-900">
            Unfilled Placeholders ({missingFields.length})
          </h3>
        </div>
        <span className="text-xs text-amber-900">
          Enter missing values below or upload the corresponding source document
        </span>
      </div>

      <div className="divide-y divide-slate-200 max-h-80 overflow-y-auto">
        {missingFields.map((item) => (
          <div
            key={item.code}
            className="px-4 py-2.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3 hover:bg-slate-50"
          >
            <div className="min-w-0">
              <div className="flex items-center gap-2">
                <span className="font-mono-tabular text-xs font-semibold text-amber-800">
                  {item.code}
                </span>
                <span className="text-sm font-medium text-slate-900">{item.label}</span>
              </div>
              <p className="text-xs text-slate-500 mt-0.5">{item.reason}</p>
            </div>

            <div className="sm:w-56 shrink-0">
              <input
                type="text"
                value={overrides[item.code] ?? ''}
                onChange={(e) => onOverrideChange(item.code, e.target.value)}
                placeholder={`Enter value for ${item.code}`}
                className="w-full px-3 py-1.5 text-sm font-mono-tabular bg-white border border-slate-300 rounded-md focus:outline-none focus:border-emerald-700"
              />
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};
