import React, { useState } from 'react';
import { CheckCircle2, FileCode2, FileUp, Trash2 } from 'lucide-react';
import type { MemoDocType, MemoTemplateRecord } from '../types/memo';

const TEMPLATE_DOCS: Array<{
  docType: MemoDocType;
  label: string;
  subtitle: string;
}> = [
  {
    docType: 'ic',
    label: 'IC Memo',
    subtitle: 'Investment Committee default .docx template',
  },
  {
    docType: 'mec',
    label: 'MEC Memo',
    subtitle: 'Management Executive Committee default .docx template',
  },
  {
    docType: 'board',
    label: 'Board Memo',
    subtitle: 'Board of Directors default .docx template',
  },
  {
    docType: 'sanction',
    label: 'Sanction',
    subtitle: 'Head Office Sanction Letter default .docx template',
  },
];

interface TemplateManagerPanelProps {
  templates: Record<MemoDocType, MemoTemplateRecord | null>;
  onUploadTemplate: (docType: MemoDocType, file: File) => Promise<void>;
  onRemoveTemplate: (docType: MemoDocType) => Promise<void>;
  singleDocType?: MemoDocType;
}

export const TemplateManagerPanel: React.FC<TemplateManagerPanelProps> = ({
  templates,
  onUploadTemplate,
  onRemoveTemplate,
  singleDocType,
}) => {
  const [uploadingType, setUploadingType] = useState<MemoDocType | null>(null);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  const handleFileChange = async (docType: MemoDocType, file: File | undefined) => {
    if (!file) return;
    setErrorMsg(null);
    setUploadingType(docType);
    try {
      await onUploadTemplate(docType, file);
    } catch (err) {
      setErrorMsg(err instanceof Error ? err.message : 'Failed to load .docx template.');
    } finally {
      setUploadingType(null);
    }
  };

  if (singleDocType) {
    const info = TEMPLATE_DOCS.find((d) => d.docType === singleDocType)!;
    const tpl = templates[singleDocType];
    const isBusy = uploadingType === singleDocType;

    return (
      <div className="bg-white border border-slate-200 rounded-lg p-4 space-y-2.5">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-start gap-3 min-w-0">
            <div className="p-2 rounded-md bg-slate-100 border border-slate-200 text-emerald-800 shrink-0">
              <FileCode2 className="w-4 h-4" />
            </div>
            <div className="min-w-0">
              <div className="flex items-center gap-2 flex-wrap">
                <h3 className="text-sm font-semibold text-slate-900">
                  {info.label} Default Format (.docx)
                </h3>
                {tpl && (
                  <span className="inline-flex items-center gap-1 text-xs font-semibold text-emerald-700">
                    <CheckCircle2 className="w-3.5 h-3.5" />
                    <span>Template Active</span>
                  </span>
                )}
              </div>
              {tpl ? (
                <p className="text-xs text-slate-600 mt-0.5 truncate">
                  Using <span className="font-medium text-slate-900">{tpl.fileName}</span> (
                  {(tpl.fileSize / 1024).toFixed(1)} KB
                  {tpl.detectedTokens.length > 0
                    ? ` · ${tpl.detectedTokens.length} placeholder tokens`
                    : ' · Smart label & table mapping'}
                  ) — exact margins, fonts, tables, headers &amp; footers preserved on export.
                </p>
              ) : (
                <p className="text-xs text-slate-500 mt-0.5">
                  {singleDocType === 'ic'
                    ? 'Optional: Upload your custom IC Memo .docx format to generate outputs using that exact layout instead of the built-in default.'
                    : `Optional: Built-in ${info.label} format is active. Upload a custom ${info.label} .docx format anytime to override layout while preserving exact formatting.`}
                </p>
              )}
            </div>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <label className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-emerald-900 bg-emerald-50 hover:bg-emerald-100 border border-emerald-300 rounded-md cursor-pointer transition-colors whitespace-nowrap">
              <FileUp className="w-3.5 h-3.5" />
              <span>{isBusy ? 'Loading...' : tpl ? 'Replace Format' : 'Upload Format (.docx)'}</span>
              <input
                type="file"
                accept=".docx"
                disabled={isBusy}
                className="hidden"
                onChange={(e) => {
                  const file = e.target.files?.[0];
                  void handleFileChange(singleDocType, file);
                  e.target.value = '';
                }}
              />
            </label>

            {tpl && (
              <button
                type="button"
                onClick={() => void onRemoveTemplate(singleDocType)}
                title="Remove uploaded template"
                className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:text-red-700 bg-white hover:bg-red-50 border border-slate-300 rounded-md cursor-pointer transition-colors"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>{singleDocType === 'ic' ? 'Use Built-in' : 'Remove'}</span>
              </button>
            )}
          </div>
        </div>

        {errorMsg && (
          <p className="text-xs text-red-600 font-medium">{errorMsg}</p>
        )}
      </div>
    );
  }

  return (
    <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
      <div className="px-5 py-4 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3">
        <div>
          <h3 className="font-display text-base font-semibold text-slate-900">
            Default Format Templates (.docx)
          </h3>
          <p className="text-xs text-slate-500 mt-0.5">
            Upload standard Word (.docx) formats for IC Memo, MEC Memo, Board Memo, and Sanction. Exported documents preserve exact margins, orientation, fonts, tables, headers, footers, and page breaks.
          </p>
        </div>
        <span className="font-mono-tabular text-xs font-semibold text-slate-700">
          {TEMPLATE_DOCS.filter((d) => templates[d.docType] !== null).length}/4 uploaded
        </span>
      </div>

      {errorMsg && (
        <div className="px-5 py-2.5 bg-red-50 border-b border-red-200 text-xs text-red-700 font-medium">
          {errorMsg}
        </div>
      )}

      <div className="p-4 grid grid-cols-1 md:grid-cols-2 gap-3">
        {TEMPLATE_DOCS.map((doc) => {
          const tpl = templates[doc.docType];
          const isBusy = uploadingType === doc.docType;

          return (
            <div
              key={doc.docType}
              className="border border-slate-200 rounded-lg p-3.5 flex items-start justify-between gap-3 bg-slate-50/50"
            >
              <div className="min-w-0 flex-1">
                <div className="flex items-center gap-2">
                  <span className="text-sm font-semibold text-slate-900">{doc.label}</span>
                  {tpl ? (
                    <span className="inline-flex items-center gap-1 text-xs font-semibold text-emerald-700">
                      <CheckCircle2 className="w-3.5 h-3.5" />
                      <span>Active</span>
                    </span>
                  ) : (
                    <span className="text-xs text-slate-400">
                      Default built-in
                    </span>
                  )}
                </div>
                <p className="text-xs text-slate-500 mt-0.5 truncate">
                  {tpl
                    ? `${tpl.fileName} (${(tpl.fileSize / 1024).toFixed(1)} KB)`
                    : doc.subtitle}
                </p>
              </div>

              <div className="flex items-center gap-1.5 shrink-0">
                <label className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-100 border border-slate-300 rounded-md cursor-pointer transition-colors whitespace-nowrap">
                  <FileUp className="w-3.5 h-3.5 text-emerald-800" />
                  <span>{isBusy ? 'Reading...' : tpl ? 'Replace' : 'Upload'}</span>
                  <input
                    type="file"
                    accept=".docx"
                    disabled={isBusy}
                    className="hidden"
                    onChange={(e) => {
                      const file = e.target.files?.[0];
                      void handleFileChange(doc.docType, file);
                      e.target.value = '';
                    }}
                  />
                </label>

                {tpl && (
                  <button
                    type="button"
                    onClick={() => void onRemoveTemplate(doc.docType)}
                    title={`Remove ${doc.label} template`}
                    className="p-1.5 text-slate-500 hover:text-red-700 bg-white hover:bg-red-50 border border-slate-300 rounded-md cursor-pointer transition-colors"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
};
