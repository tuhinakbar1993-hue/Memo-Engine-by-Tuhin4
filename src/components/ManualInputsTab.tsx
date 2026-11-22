import React, { useState } from 'react';
import {
  Check,
  ChevronDown,
  ChevronRight,
  FileSpreadsheet,
  FileUp,
  FolderOpen,
  Loader2,
  RefreshCw,
  RotateCcw,
  Trash2,
  X,
} from 'lucide-react';
import { IC_PLACEHOLDER_SPECS } from '../engines/icMemoEngine';
import type {
  DocumentKey,
  DocumentUploadMeta,
  ExtractedDocumentsState,
  ManualInputsData,
  MemoDocType,
  MemoTemplateRecord,
  ResolvedICMemo,
} from '../types/memo';
import { TemplateManagerPanel } from './TemplateManagerPanel';
import { ValidationPanel } from './ValidationPanel';

interface ManualInputsTabProps {
  docMetas: Record<DocumentKey, DocumentUploadMeta>;
  extractedDocs: ExtractedDocumentsState;
  manualInputs: ManualInputsData;
  overrides: Record<string, string>;
  resolvedMemo: ResolvedICMemo;
  onFileUpload: (key: DocumentKey, file: File) => void;
  onBatchUpload: (files: FileList | File[]) => Promise<string[]>;
  onManualChange: (field: keyof ManualInputsData, value: string) => void;
  onOverrideChange: (code: string, value: string) => void;
  onResetOverrides: () => void;
  onLoadSampleDocuments: () => void;
  loadingSamples: boolean;
  batchProgress: { active: boolean; total: number; completed: number };
  onGoToICMemo: () => void;
  templates: Record<MemoDocType, MemoTemplateRecord | null>;
  onUploadTemplate: (docType: MemoDocType, file: File) => Promise<void>;
  onRemoveTemplate: (docType: MemoDocType) => Promise<void>;
  onApplyAllChanges: () => void;
  hasPendingChanges: boolean;
  onClearAllData: () => void;
}

const DOC_ORDER: DocumentKey[] = [
  'branchProposal',
  'outstanding',
  'partyPerformance',
  'overdue',
  'accountStatement',
  'cibReport',
  'costOfFund',
];

export const ManualInputsTab: React.FC<ManualInputsTabProps> = ({
  docMetas,
  manualInputs,
  overrides,
  resolvedMemo,
  onFileUpload,
  onBatchUpload,
  onManualChange,
  onOverrideChange,
  onResetOverrides,
  onLoadSampleDocuments,
  loadingSamples,
  batchProgress,
  onGoToICMemo,
  templates,
  onUploadTemplate,
  onRemoveTemplate,
  onApplyAllChanges,
  hasPendingChanges,
  onClearAllData,
}) => {
  const [showAllFields, setShowAllFields] = useState(true);
  const [searchFilter, setSearchFilter] = useState('');
  const [isDraggingBatch, setIsDraggingBatch] = useState(false);
  const [batchNotice, setBatchNotice] = useState<string | null>(null);

  const handleBatchFiles = async (files: FileList | null) => {
    if (!files || files.length === 0) return;
    setBatchNotice(null);
    const unmatched = await onBatchUpload(files);
    if (unmatched.length > 0) {
      setBatchNotice(
        `Could not auto-match: ${unmatched.join(', ')}. Use the individual row buttons below.`
      );
    }
  };

  const uploadedCount = DOC_ORDER.filter((k) => docMetas[k].uploaded).length;
  const anyParsing = DOC_ORDER.some((k) => docMetas[k].parsing);
  const progressPct =
    batchProgress.total > 0
      ? Math.round((batchProgress.completed / batchProgress.total) * 100)
      : 0;
  const isCalc = (code: string) => code === '#32#' || code === '#42#' || code === '#44#';

  const filteredSpecs = IC_PLACEHOLDER_SPECS.filter((spec) => {
    if (!searchFilter.trim()) return true;
    const q = searchFilter.toLowerCase();
    return (
      spec.code.toLowerCase().includes(q) ||
      spec.label.toLowerCase().includes(q) ||
      spec.source.toLowerCase().includes(q)
    );
  });

  return (
    <div className="max-w-7xl mx-auto space-y-6">
      {/* Top Action & Validation Summary */}
      <div className="bg-white border border-slate-200 rounded-lg p-5 flex flex-col lg:flex-row lg:items-center justify-between gap-4">
        <div>
          <h2 className="font-display text-lg font-bold text-slate-900">
            Document Ingestion &amp; Manual Inputs
          </h2>
          <p className="text-sm text-slate-600 mt-0.5">
            Upload the seven banking documents, enter proposal terms, and review extracted Million BDT figures before generating the memo.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-3 shrink-0">
          <button
            type="button"
            onClick={onLoadSampleDocuments}
            disabled={loadingSamples}
            className="inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded-lg transition-colors cursor-pointer disabled:opacity-60 whitespace-nowrap shrink-0"
          >
            {loadingSamples ? (
              <Loader2 className="w-4 h-4 animate-spin" />
            ) : (
              <FolderOpen className="w-4 h-4" />
            )}
            <span>{loadingSamples ? 'Parsing Samples...' : 'Load Attached Sample Case'}</span>
          </button>

          <button
            type="button"
            onClick={onApplyAllChanges}
            className={`inline-flex items-center gap-2 px-4 py-2 text-sm font-semibold rounded-lg border transition-colors cursor-pointer whitespace-nowrap shrink-0 ${
              hasPendingChanges
                ? 'text-amber-950 bg-amber-100 hover:bg-amber-200 border-amber-300'
                : 'text-slate-700 bg-white hover:bg-slate-100 border-slate-300'
            }`}
          >
            <RefreshCw className="w-4 h-4" />
            <span>Apply All Changes</span>
          </button>

          <button
            type="button"
            onClick={onClearAllData}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 text-sm font-medium text-red-700 bg-red-50 hover:bg-red-100 border border-red-200 rounded-lg transition-colors cursor-pointer whitespace-nowrap shrink-0"
          >
            <Trash2 className="w-4 h-4" />
            <span>Clear All Data</span>
          </button>

          <button
            type="button"
            onClick={onGoToICMemo}
            className="inline-flex items-center gap-2 px-4 py-2 text-sm font-semibold text-white bg-emerald-800 hover:bg-emerald-900 rounded-lg transition-colors cursor-pointer whitespace-nowrap shrink-0"
          >
            <span>Open IC Memo Preview</span>
          </button>
        </div>
      </div>

      {/* Main 2-Column Workspace */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
        {/* Left Column (5 cols): 7 Document Upload Checklist */}
        <div className="lg:col-span-5 bg-white border border-slate-200 rounded-lg overflow-hidden">
          <div className="px-5 py-4 border-b border-slate-200 flex items-center justify-between">
            <div>
              <h3 className="font-display text-base font-semibold text-slate-900">
                Source Documents
              </h3>
              <p className="text-xs text-slate-500 mt-0.5">
                {uploadedCount} of 7 documents uploaded
              </p>
            </div>
            <span className="font-mono-tabular text-xs font-semibold text-slate-700">
              {uploadedCount}/7
            </span>
          </div>

          {(batchProgress.active || anyParsing) && (
            <div className="px-5 py-2.5 bg-emerald-50/70 border-b border-emerald-200 space-y-1.5">
              <div className="flex items-center justify-between text-xs font-medium text-emerald-900">
                <span className="inline-flex items-center gap-1.5">
                  <Loader2 className="w-3.5 h-3.5 animate-spin text-emerald-800" />
                  <span>
                    {batchProgress.active
                      ? `Processing documents (${batchProgress.completed} of ${batchProgress.total} completed)...`
                      : 'Processing uploaded document...'}
                  </span>
                </span>
                {batchProgress.active && (
                  <span className="font-mono-tabular font-semibold">{progressPct}%</span>
                )}
              </div>
              {batchProgress.active && (
                <div className="w-full h-1.5 bg-emerald-200/70 rounded-full overflow-hidden">
                  <div
                    className="h-full bg-emerald-700 transition-all duration-200"
                    style={{ width: `${Math.max(8, progressPct)}%` }}
                  />
                </div>
              )}
            </div>
          )}

          {/* Quick Batch Drop / Multi-Select Upload */}
          <div className="p-3 bg-slate-50 border-b border-slate-200">
            <label
              onDragOver={(e) => {
                e.preventDefault();
                setIsDraggingBatch(true);
              }}
              onDragLeave={() => setIsDraggingBatch(false)}
              onDrop={(e) => {
                e.preventDefault();
                setIsDraggingBatch(false);
                void handleBatchFiles(e.dataTransfer.files);
              }}
              className={`flex items-center justify-between gap-3 px-3.5 py-2.5 rounded-md border border-dashed transition-colors cursor-pointer ${
                isDraggingBatch
                  ? 'border-emerald-700 bg-emerald-50/60'
                  : 'border-slate-300 bg-white hover:border-emerald-700'
              }`}
            >
              <div className="min-w-0">
                <p className="text-xs font-semibold text-slate-800">
                  Quick Batch Upload (Drop or Select All 7 Files)
                </p>
                <p className="text-[11px] text-slate-500 truncate">
                  Auto-routes and parses all files concurrently
                </p>
              </div>
              <span className="inline-flex items-center gap-1.5 px-2.5 py-1 text-xs font-semibold text-emerald-800 bg-emerald-50 border border-emerald-200 rounded shrink-0 whitespace-nowrap">
                <FileUp className="w-3.5 h-3.5" />
                <span>Select Files</span>
              </span>
              <input
                type="file"
                multiple
                accept=".pdf,.docx,.xlsx,.xls,.csv,image/*"
                className="hidden"
                onChange={(e) => {
                  void handleBatchFiles(e.target.files);
                  e.target.value = '';
                }}
              />
            </label>
            {batchNotice && (
              <p className="mt-1.5 text-xs text-amber-800 font-medium">{batchNotice}</p>
            )}
          </div>

          <div className="divide-y divide-slate-200">
            {DOC_ORDER.map((key, index) => {
              const meta = docMetas[key];
              const isExcel = key === 'costOfFund';

              return (
                <div key={key} className="p-4 flex items-start justify-between gap-3 hover:bg-slate-50/70">
                  <div className="min-w-0 flex-1">
                    <div className="flex items-center gap-2">
                      <span className="font-mono-tabular text-xs text-slate-400">
                        0{index + 1}.
                      </span>
                      <span className="text-sm font-semibold text-slate-900 truncate">
                        {meta.label}
                      </span>
                      {meta.parsing ? (
                        <span className="inline-flex items-center gap-1 text-xs font-medium text-emerald-800">
                          <Loader2 className="w-3.5 h-3.5 animate-spin shrink-0" />
                          <span className="truncate">{meta.stage || 'Reading...'}</span>
                        </span>
                      ) : meta.uploaded ? (
                        <span
                          className="inline-flex items-center gap-1 text-xs font-semibold text-emerald-700"
                          title="Uploaded and parsed"
                        >
                          <Check className="w-3.5 h-3.5" />
                          <span>✓</span>
                        </span>
                      ) : (
                        <span
                          className="inline-flex items-center gap-1 text-xs font-semibold text-slate-400"
                          title="Not uploaded"
                        >
                          <X className="w-3.5 h-3.5" />
                          <span>✗</span>
                        </span>
                      )}
                    </div>

                    <p className="text-xs text-slate-500 mt-1 truncate">
                      {meta.uploaded || meta.parsing
                        ? meta.fileName
                        : key === 'branchProposal'
                          ? 'PDF, Word (.docx), Excel (.xlsx/.xls), or Image'
                          : key === 'costOfFund'
                            ? 'Excel (.xlsx / .xls), Word (.docx), or PDF'
                            : 'PDF, Word (.docx), or Excel (.xlsx / .xls)'}
                    </p>

                    {meta.error && (
                      <p className="text-xs text-red-600 mt-1 font-medium">{meta.error}</p>
                    )}
                  </div>

                  <label className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-100 border border-slate-300 rounded-md cursor-pointer transition-colors shrink-0 whitespace-nowrap">
                    {isExcel ? (
                      <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-800" />
                    ) : (
                      <FileUp className="w-3.5 h-3.5 text-emerald-800" />
                    )}
                    <span>{meta.uploaded ? 'Replace' : 'Upload'}</span>
                    <input
                      type="file"
                      accept={meta.acceptedFormats}
                      className="hidden"
                      onChange={(e) => {
                        const file = e.target.files?.[0];
                        if (file) onFileUpload(key, file);
                        e.target.value = '';
                      }}
                    />
                  </label>
                </div>
              );
            })}
          </div>
        </div>

        {/* Right Column (7 cols): Manual Inputs Form + Validation */}
        <div className="lg:col-span-7 space-y-6">
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            <div className="px-5 py-4 border-b border-slate-200 flex items-center justify-between">
              <div>
                <h3 className="font-display text-base font-semibold text-slate-900">
                  Manual Inputs (Million BDT)
                </h3>
                <p className="text-xs text-slate-500 mt-0.5">
                  Values typed here are used directly in the IC Memo without unit conversion.
                </p>
              </div>
            </div>

            <div className="p-5 grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Proposal Date <span className="font-mono-tabular text-slate-400">(#1#)</span>
                </label>
                <input
                  type="text"
                  value={manualInputs.proposalDate}
                  onChange={(e) => onManualChange('proposalDate', e.target.value)}
                  placeholder="e.g. 18.11.2026"
                  className="w-full px-3 py-2 text-sm font-mono-tabular bg-white border border-slate-300 rounded-md focus:outline-none focus:border-emerald-700"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Waiver Amount (Million BDT) <span className="font-mono-tabular text-slate-400">(#4#)</span>
                </label>
                <input
                  type="text"
                  value={manualInputs.waiverAmount}
                  onChange={(e) => onManualChange('waiverAmount', e.target.value)}
                  placeholder="e.g. 0.06"
                  className="w-full px-3 py-2 text-sm font-mono-tabular bg-white border border-slate-300 rounded-md focus:outline-none focus:border-emerald-700"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Settlement Amount (Million BDT) <span className="font-mono-tabular text-slate-400">(#27#)</span>
                </label>
                <input
                  type="text"
                  value={manualInputs.settlementAmount}
                  onChange={(e) => onManualChange('settlementAmount', e.target.value)}
                  placeholder="e.g. 0.22"
                  className="w-full px-3 py-2 text-sm font-mono-tabular bg-white border border-slate-300 rounded-md focus:outline-none focus:border-emerald-700"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Settlement Period <span className="font-mono-tabular text-slate-400">(#28#)</span>
                </label>
                <input
                  type="text"
                  value={manualInputs.settlementPeriod}
                  onChange={(e) => onManualChange('settlementPeriod', e.target.value)}
                  placeholder="e.g. 30.12.2026"
                  className="w-full px-3 py-2 text-sm font-mono-tabular bg-white border border-slate-300 rounded-md focus:outline-none focus:border-emerald-700"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Waiver Recommended by Client &amp; Branch <span className="font-mono-tabular text-slate-400">(#45#)</span>
                </label>
                <input
                  type="text"
                  value={manualInputs.waiverRecClientBranch}
                  onChange={(e) => onManualChange('waiverRecClientBranch', e.target.value)}
                  placeholder="e.g. 0.06"
                  className="w-full px-3 py-2 text-sm font-mono-tabular bg-white border border-slate-300 rounded-md focus:outline-none focus:border-emerald-700"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Waiver Recommended by H/O <span className="font-mono-tabular text-slate-400">(#47#)</span>
                </label>
                <input
                  type="text"
                  value={manualInputs.waiverRecHO}
                  onChange={(e) => onManualChange('waiverRecHO', e.target.value)}
                  placeholder="e.g. 0.06"
                  className="w-full px-3 py-2 text-sm font-mono-tabular bg-white border border-slate-300 rounded-md focus:outline-none focus:border-emerald-700"
                />
              </div>

              <div className="sm:col-span-2">
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Source of Repayment <span className="font-mono-tabular text-slate-400">(Section 8)</span>
                </label>
                <textarea
                  rows={2}
                  value={manualInputs.sourceOfRepayment}
                  onChange={(e) => onManualChange('sourceOfRepayment', e.target.value)}
                  placeholder="e.g. Own source / family support and liquidation of personal assets."
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-md focus:outline-none focus:border-emerald-700"
                />
              </div>
            </div>
          </div>

          {/* Unfilled Placeholders Validation Panel */}
          <ValidationPanel
            missingFields={resolvedMemo.missingFields}
            overrides={overrides}
            onOverrideChange={onOverrideChange}
          />
        </div>
      </div>

      {/* Default Format Templates Upload Panel (IC Memo, MEC Memo, Board Memo, Sanction) */}
      <TemplateManagerPanel
        templates={templates}
        onUploadTemplate={onUploadTemplate}
        onRemoveTemplate={onRemoveTemplate}
      />

      {/* Review & Correct All Extracted Placeholders Table */}
      <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
        <div className="px-5 py-3.5 border-b border-slate-200 flex flex-wrap items-center justify-between gap-4">
          <button
            type="button"
            onClick={() => setShowAllFields((v) => !v)}
            className="flex items-center gap-2 text-left cursor-pointer"
          >
            {showAllFields ? (
              <ChevronDown className="w-4 h-4 text-slate-600" />
            ) : (
              <ChevronRight className="w-4 h-4 text-slate-600" />
            )}
            <div>
              <h3 className="font-display text-base font-semibold text-slate-900">
                Extracted Fields &amp; Unit Conversion Review
              </h3>
              <p className="text-xs text-slate-500">
                Inspect or override any extracted value (figures from non-proposal sources are automatically converted from BDT to Million BDT).
              </p>
            </div>
          </button>

          {showAllFields && (
            <div className="flex items-center gap-3">
              <input
                type="text"
                value={searchFilter}
                onChange={(e) => setSearchFilter(e.target.value)}
                placeholder="Filter by #code or field..."
                className="px-3 py-1.5 text-xs bg-white border border-slate-300 rounded-md focus:outline-none focus:border-emerald-700 w-52"
              />
              {Object.keys(overrides).length > 0 && (
                <button
                  type="button"
                  onClick={onResetOverrides}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-md cursor-pointer whitespace-nowrap"
                >
                  <RotateCcw className="w-3.5 h-3.5" />
                  <span>Reset Overrides</span>
                </button>
              )}
            </div>
          )}
        </div>

        {showAllFields && (
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-600">
                  <th className="py-2.5 px-4 w-24">Code</th>
                  <th className="py-2.5 px-4">IC Memo Field</th>
                  <th className="py-2.5 px-4">Authoritative Source</th>
                  <th className="py-2.5 px-4 w-64">Resolved Value (Editable)</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 text-sm">
                {filteredSpecs.map((spec) => {
                  const val = resolvedMemo.placeholders[spec.code] ?? '';
                  const isOverridden = overrides[spec.code] !== undefined && overrides[spec.code] !== '';
                  const missing = !val;

                  return (
                    <tr key={spec.code} className="hover:bg-slate-50/70">
                      <td className="py-2 px-4 font-mono-tabular text-xs font-semibold text-slate-700">
                        {spec.code}
                      </td>
                      <td className="py-2 px-4 text-slate-900 font-medium">{spec.label}</td>
                      <td className="py-2 px-4 text-xs text-slate-500">{spec.source}</td>
                      <td className="py-2 px-4">
                        <div className="flex items-center gap-2">
                          <input
                            type="text"
                            value={overrides[spec.code] !== undefined ? overrides[spec.code] : val}
                            onChange={(e) => onOverrideChange(spec.code, e.target.value)}
                            placeholder={isCalc(spec.code) ? 'Auto-calculated' : 'Missing'}
                            className={`w-full px-2.5 py-1 text-sm font-mono-tabular rounded border focus:outline-none focus:border-emerald-700 ${
                              missing
                                ? 'border-amber-300 bg-amber-50/40 text-slate-900'
                                : isOverridden
                                  ? 'border-emerald-400 bg-emerald-50/30 text-slate-900'
                                  : 'border-slate-200 bg-white text-slate-900'
                            }`}
                          />
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
};
