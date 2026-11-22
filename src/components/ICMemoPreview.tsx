import React, { useEffect, useState } from 'react';
import {
  AlertTriangle,
  CheckCircle2,
  Download,
  Loader2,
  Pencil,
  Plus,
  RefreshCw,
  RotateCcw,
  Sparkles,
  Trash2,
  X,
} from 'lucide-react';
import type {
  AiSectionsData,
  BatchScrutinyState,
  MemoDocType,
  MemoTemplateRecord,
  ResolvedICMemo,
  ScrutinySectionKey,
  Section2AiData,
  Section3AiData,
  Section4AiData,
  Section5AiData,
  Section6AiData,
  Section7AiData,
  Section9AiData,
  Section10AiData,
  Section10Particulars,
  Section11AiData,
} from '../types/memo';
import { TemplateManagerPanel } from './TemplateManagerPanel';
import { ValidationPanel } from './ValidationPanel';

const SECTION_LABELS: Record<ScrutinySectionKey, string> = {
  justification: 'Justification of The Proposed Facility',
  sec1NatureOfBusiness: '1. Nature of Business',
  sec2LiabilityOurBank: '2. Liability Position with Our Bank',
  sec3LiabilityOtherUnits: '3. Liability with Other Units / Banks',
  sec4CibStatus: '4. CIB Status',
  sec5LegalAction: '5. Legal Action Status',
  sec6AuditCompliance: '6. Audit, Compliance & Regulatory Status',
  sec7Security: '7. Security',
  sec8SourceOfRepayment: '8. Source of Repayment',
  sec9HoSanctions: '9. Sequence of Head Office Sanction(s)',
  sec10ClientParticulars: '10. Brief Particulars of The Client',
  sec11Financials: '11. Financials of The Client',
  clientPrayer: "12. Contents of the Client’s Prayer",
  branchVisit: '12. Branch Visit Report',
  sec13BranchRecommendation: '13. Recommendation of The Branch',
  sec14IccdRecommendation: '14. Recommendation of ICCD',
  samdObservation: '15. Observation and Recommendation of SAMD',
};

interface ICMemoPreviewProps {
  memo: ResolvedICMemo;
  aiSections: AiSectionsData;
  overrides: Record<string, string>;
  onOverrideChange: (code: string, value: string) => void;
  onScrutinizeSection: (section: ScrutinySectionKey, forceOverwrite?: boolean) => Promise<void>;
  onRetrySection: (section: ScrutinySectionKey) => Promise<void>;
  onScrutinizeAll: () => Promise<void>;
  onDismissBatchSummary: () => void;
  batchScrutiny: BatchScrutinyState;
  onUpdateAiSections: (
    updater: (prev: AiSectionsData) => AiSectionsData,
    markManualSection?: ScrutinySectionKey
  ) => void;
  manualEditedSections: Partial<Record<ScrutinySectionKey, boolean>>;
  sourceOfRepaymentManual: string;
  onManualSourceOfRepaymentChange: (value: string) => void;
  scrutinizingSection: ScrutinySectionKey | null;
  hasBranchProposal: boolean;
  aiError: string | null;
  onDownloadDocx: () => Promise<void>;
  onDownloadPdf: () => Promise<void>;
  downloading: boolean;
  downloadingPdf: boolean;
  templates: Record<MemoDocType, MemoTemplateRecord | null>;
  onUploadTemplate: (docType: MemoDocType, file: File) => Promise<void>;
  onRemoveTemplate: (docType: MemoDocType) => Promise<void>;
  onApplyAllChanges: () => void;
  hasPendingChanges: boolean;
  applySignal: number;
}

const Token: React.FC<{ ph: Record<string, string>; code: string; bold?: boolean }> = ({
  ph,
  code,
  bold = true,
}) => {
  const val = ph[code];
  const isMissing = !val || val.trim() === '';
  if (isMissing) {
    return (
      <span className="bg-amber-100 text-amber-900 px-1 rounded font-mono text-[9pt] font-semibold">
        {code}
      </span>
    );
  }
  return <span className={bold ? 'font-bold' : ''}>{val}</span>;
};

interface SectionToolbarProps {
  sectionKey: ScrutinySectionKey;
  editing: boolean;
  onToggleEdit: () => void;
  scrutinizingSection: ScrutinySectionKey | null;
  batchRunning: boolean;
  isRetrying: boolean;
  failedError: string | null;
  insufficientReason: string | null;
  isManualProtected: boolean;
  confirmOverwrite: boolean;
  onRequestScrutiny: () => void;
  onRetry: () => void;
  onConfirmOverwrite: () => void;
  onCancelOverwrite: () => void;
}

const SectionActionToolbar: React.FC<SectionToolbarProps> = ({
  sectionKey,
  editing,
  onToggleEdit,
  scrutinizingSection,
  batchRunning,
  isRetrying,
  failedError,
  insufficientReason,
  isManualProtected,
  confirmOverwrite,
  onRequestScrutiny,
  onRetry,
  onConfirmOverwrite,
  onCancelOverwrite,
}) => {
  const isRunning = scrutinizingSection === sectionKey || isRetrying;
  const anyRunning = scrutinizingSection !== null || batchRunning;

  return (
    <div className="font-sans space-y-1.5">
      <div className="flex items-center gap-1.5 flex-wrap justify-end">
        <button
          type="button"
          onClick={onToggleEdit}
          className={`inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-medium rounded border cursor-pointer transition-colors ${
            editing
              ? 'text-emerald-900 bg-emerald-100 hover:bg-emerald-200 border-emerald-400'
              : 'text-slate-700 bg-slate-100 hover:bg-slate-200 border-slate-300'
          }`}
        >
          <Pencil className="w-3 h-3" />
          <span>{editing ? 'Done Editing' : 'Edit'}</span>
        </button>

        {(failedError || isRetrying) && (
          <button
            type="button"
            onClick={onRetry}
            disabled={isRetrying || batchRunning || scrutinizingSection !== null}
            title={failedError || 'Retry AI scrutiny for this section'}
            className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-semibold text-white bg-amber-700 hover:bg-amber-800 rounded cursor-pointer disabled:opacity-60 transition-colors whitespace-nowrap"
          >
            {isRetrying ? (
              <Loader2 className="w-3 h-3 animate-spin" />
            ) : (
              <RotateCcw className="w-3 h-3" />
            )}
            <span>{isRetrying ? 'Retrying...' : 'Retry'}</span>
          </button>
        )}

        <button
          type="button"
          onClick={onRequestScrutiny}
          disabled={anyRunning || isRetrying}
          title={
            isManualProtected
              ? 'This section has manual edits — clicking will ask before overwriting'
              : 'Scrutinize uploaded Branch Proposal and populate this section with AI'
          }
          className="inline-flex items-center gap-1.5 px-2.5 py-1 text-[11px] font-semibold text-white bg-emerald-800 hover:bg-emerald-900 rounded cursor-pointer disabled:opacity-60 transition-colors whitespace-nowrap"
        >
          {isRunning ? (
            <Loader2 className="w-3.5 h-3.5 animate-spin" />
          ) : (
            <Sparkles className="w-3.5 h-3.5" />
          )}
          <span>{isRunning ? 'Scrutinizing...' : 'Scrutinize and Fill Up Through AI'}</span>
        </button>
      </div>

      {failedError && !isRetrying && (
        <div className="flex flex-wrap items-center justify-between gap-2 bg-red-50 border border-red-200 rounded px-2.5 py-1 text-[11px] text-red-800">
          <span>{failedError}</span>
        </div>
      )}

      {!failedError && insufficientReason && (
        <div className="flex flex-wrap items-center justify-between gap-2 bg-amber-50 border border-amber-200 rounded px-2.5 py-1 text-[11px] text-amber-900">
          <span>{insufficientReason}</span>
        </div>
      )}

      {confirmOverwrite && (
        <div className="flex flex-wrap items-center justify-between gap-2 bg-amber-50 border border-amber-300 rounded px-2.5 py-1.5 text-[11px] text-amber-950">
          <div className="flex items-center gap-1.5 font-medium">
            <AlertTriangle className="w-3.5 h-3.5 text-amber-700 shrink-0" />
            <span>
              This section contains manually entered or edited content. Overwrite it with AI scrutiny?
            </span>
          </div>
          <div className="flex items-center gap-1.5">
            <button
              type="button"
              onClick={onConfirmOverwrite}
              disabled={anyRunning}
              className="px-2 py-0.5 font-semibold text-white bg-amber-700 hover:bg-amber-800 rounded cursor-pointer disabled:opacity-60"
            >
              Yes, Overwrite
            </button>
            <button
              type="button"
              onClick={onCancelOverwrite}
              className="px-2 py-0.5 font-medium text-slate-700 bg-white hover:bg-slate-100 border border-slate-300 rounded cursor-pointer"
            >
              Cancel
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

export const ICMemoPreview: React.FC<ICMemoPreviewProps> = ({
  memo,
  aiSections,
  overrides,
  onOverrideChange,
  onScrutinizeSection,
  onRetrySection,
  onScrutinizeAll,
  onDismissBatchSummary,
  batchScrutiny,
  onUpdateAiSections,
  manualEditedSections,
  sourceOfRepaymentManual,
  onManualSourceOfRepaymentChange,
  scrutinizingSection,
  hasBranchProposal,
  aiError,
  onDownloadDocx,
  onDownloadPdf,
  downloading,
  downloadingPdf,
  templates,
  onUploadTemplate,
  onRemoveTemplate,
  onApplyAllChanges,
  hasPendingChanges,
  applySignal,
}) => {
  const ph = memo.placeholders;
  const bp = memo.branchProposal;

  const [editingSections, setEditingSections] = useState<Partial<Record<ScrutinySectionKey, boolean>>>({});
  const [confirmOverwriteSection, setConfirmOverwriteSection] = useState<ScrutinySectionKey | null>(null);

  useEffect(() => {
    if (applySignal > 0) {
      setEditingSections({});
      setConfirmOverwriteSection(null);
    }
  }, [applySignal]);

  const toggleEditing = (key: ScrutinySectionKey) => {
    setEditingSections((prev) => ({ ...prev, [key]: !prev[key] }));
  };

  const isEditing = (key: ScrutinySectionKey) => Boolean(editingSections[key]);

  const isProtectedSection = (key: ScrutinySectionKey): boolean => {
    if (manualEditedSections[key]) return true;
    if (key === 'sec8SourceOfRepayment' && sourceOfRepaymentManual.trim() !== '') return true;
    return false;
  };

  const handleRequestScrutiny = (key: ScrutinySectionKey) => {
    if (isProtectedSection(key)) {
      setConfirmOverwriteSection(key);
      return;
    }
    setConfirmOverwriteSection(null);
    void onScrutinizeSection(key, false);
  };

  const handleConfirmOverwrite = (key: ScrutinySectionKey) => {
    setConfirmOverwriteSection(null);
    void onScrutinizeSection(key, true);
  };

  const toolbarProps = (key: ScrutinySectionKey): SectionToolbarProps => {
    const failEntry = batchScrutiny.failedSections.find((f) => f.section === key);
    const insuffEntry = batchScrutiny.insufficientSections.find((i) => i.section === key);
    return {
      sectionKey: key,
      editing: isEditing(key),
      onToggleEdit: () => toggleEditing(key),
      scrutinizingSection,
      batchRunning: batchScrutiny.running,
      isRetrying: batchScrutiny.retryingSections.includes(key),
      failedError: failEntry ? failEntry.error : null,
      insufficientReason: insuffEntry ? insuffEntry.reason : null,
      isManualProtected: isProtectedSection(key),
      confirmOverwrite: confirmOverwriteSection === key,
      onRequestScrutiny: () => handleRequestScrutiny(key),
      onRetry: () => void onRetrySection(key),
      onConfirmOverwrite: () => handleConfirmOverwrite(key),
      onCancelOverwrite: () => setConfirmOverwriteSection(null),
    };
  };

  // Default fallback structures preserving the official IC Memo layout
  const defaultCibStatuses = [
    'STD',
    'SMA',
    'SS',
    'DF',
    'BL',
    'Willful Default (WD)',
    'Write-Off',
    'Non-Funded',
    'Stay Order',
    'TOTAL',
  ];

  const rawSec3Rows =
    bp?.section3Rows && bp.section3Rows.length > 0
      ? bp.section3Rows
      : [
          { sl: 'a.', nameOfConcern: '', bankOrNbfi: '', natureOfFacility: '', limit: '', outstanding: '', overdueEol: '', expiryDate: '', clStatus: '' },
          { sl: 'b.', nameOfConcern: '', bankOrNbfi: '', natureOfFacility: '', limit: '', outstanding: '', overdueEol: '', expiryDate: '', clStatus: '' },
          { sl: 'c.', nameOfConcern: '', bankOrNbfi: '', natureOfFacility: '', limit: '', outstanding: '', overdueEol: '', expiryDate: '', clStatus: '' },
        ];
  const hasSec3SubTotal = rawSec3Rows.some((r) => /sub[-\s]*total/i.test(r.sl) || /sub[-\s]*total/i.test(r.nameOfConcern));
  const hasSec3GrandTotal = rawSec3Rows.some((r) => /grand\s*total/i.test(r.sl) || /grand\s*total/i.test(r.nameOfConcern));

  const sec4Rows =
    bp?.section4Rows && bp.section4Rows.length > 0
      ? bp.section4Rows
      : defaultCibStatuses.map((s) => ({
          status: s,
          unitBorrower: '',
          unitGuarantor: '',
          otherBorrower: '',
          otherGuarantor: '',
          remarks: '',
        }));

  const sec5Rows =
    bp?.section5Rows && bp.section5Rows.length > 0
      ? bp.section5Rows
      : [{ sl: '', typeOfCase: '', caseNo: '', filingDate: '', amount: '', previousDate: '', nextDate: '', presentStatus: '' }];

  const defaultSec6 = [
    { sl: 'a.', authority: 'IC & CD', lastInspectionDate: '', objectionStatus: '', objections: '', branchReply: '' },
    { sl: 'b.', authority: 'DBI, Bangladesh Bank', lastInspectionDate: '', objectionStatus: 'No', objections: '', branchReply: '' },
    { sl: 'c.', authority: 'DOS, Bangladesh Bank', lastInspectionDate: 'Not done', objectionStatus: 'No', objections: '', branchReply: '' },
    { sl: 'd.', authority: 'BFIU, Bangladesh Bank', lastInspectionDate: 'Not done', objectionStatus: 'No', objections: '', branchReply: '' },
    { sl: 'e', authority: 'BRD, Bangladesh Bank', lastInspectionDate: 'Not done', objectionStatus: 'No', objections: '', branchReply: '' },
    { sl: 'f', authority: 'ACC(Anti-Corruption Commission)', lastInspectionDate: 'Not done', objectionStatus: 'No', objections: '', branchReply: '' },
  ];
  const sec6Rows = bp?.section6Rows && bp.section6Rows.length > 0 ? bp.section6Rows : defaultSec6;

  const sec7Primary =
    bp?.section7PrimaryRows && bp.section7PrimaryRows.length > 0
      ? bp.section7PrimaryRows
      : [
          { particulars: 'i) Hypothecation of Stock as on ………………..', value: '' },
          { particulars: 'ii) Hypothecation of Machinery', value: '' },
          { particulars: 'iii) Assignment of Bills against Work Order', value: '' },
          { particulars: 'iv) Margin held against existing LC & BG liability as on……..', value: '' },
        ];

  const sec7Schedules =
    bp?.section7CollateralSchedules && bp.section7CollateralSchedules.length > 0
      ? bp.section7CollateralSchedules
      : [
          'Schedule-01: Full Schedule with owner name as per Mortgage Deed.',
          'Schedule-02: Full Schedule with owner name as per Mortgage Deed.',
          'Schedule-03: Full Schedule with owner name as per Mortgage Deed.',
          'Schedule-04: Full Schedule with owner name as per Mortgage Deed.',
        ];

  const sec7ValRows =
    bp?.section7ValuationRows && bp.section7ValuationRows.length > 0
      ? bp.section7ValuationRows
      : ['Schedule-01', 'Schedule-03', 'Schedule-02', 'Schedule-04', 'Building', 'Other', 'Grand Total'].map((s) => ({
          schedule: s,
          origBranchMv: '',
          origBranchFsv: '',
          origSurvMv: '',
          origSurvFsv: '',
          presBranchMv: '',
          presBranchFsv: '',
          presSurvMv: '',
          presSurvFsv: '',
          presAvMv: '',
          presAvFsv: '',
        }));

  const sec9Rows =
    bp?.section9Rows && bp.section9Rows.length > 0
      ? bp.section9Rows
      : ['a.', 'b.', 'c.', 'd.', 'e.', 'f.', 'g.'].map((sl) => ({
          sl,
          referenceNo: '',
          date: '',
          limit: '',
          remarks: '',
        }));

  const sp = bp?.section10Particulars;

  const emptyParticulars = (): Section10Particulars => ({
    nameOfClient: sp?.nameOfClient || '',
    uniqueCustomerId: sp?.uniqueCustomerId || '',
    awcdMsdAccount: sp?.awcdMsdAccount || '',
    businessTypeNature: sp?.businessTypeNature || '',
    businessManagementType: sp?.businessManagementType || '',
    clientType: sp?.clientType || '',
    keyPersonPhone: sp?.keyPersonPhone || '',
    address: sp?.address || '',
    officeBusinessAddress: sp?.officeBusinessAddress || '',
    factoryAddress: sp?.factoryAddress || '',
    purposeOfInvestment: sp?.purposeOfInvestment || '',
    limitOfInvestment: sp?.limitOfInvestment || '',
    modeOfInvestment: sp?.modeOfInvestment || '',
    lastVisitDate: sp?.lastVisitDate || '',
    stockPosition: sp?.stockPosition || '',
    fixedAsset: sp?.fixedAsset || '',
    businessNetWorth: sp?.businessNetWorth || '',
    insuranceCoverage: sp?.insuranceCoverage || '',
    tradeLicenseValidity: sp?.tradeLicenseValidity || '',
    otherRegulatoryDocuments: sp?.otherRegulatoryDocuments || '',
    dateOfOriginalInvestment: sp?.dateOfOriginalInvestment || '',
    firstDisbursementManager: sp?.firstDisbursementManager || '',
    firstDateOfClassification: sp?.firstDateOfClassification || '',
  });

  const dirRows =
    bp?.section10Directors && bp.section10Directors.length > 0
      ? bp.section10Directors
      : [
          { name: '', age: '', education: '', designation: '', sharePct: '', experience: '', netWorth: '' },
          { name: '', age: '', education: '', designation: '', sharePct: '', experience: '', netWorth: '' },
        ];

  const bankDirRows =
    bp?.section10BankDirectors && bp.section10BankDirectors.length > 0
      ? bp.section10BankDirectors
      : [
          { name: '', bankNbfi: '', designation: '', sharePct: '', numShares: '' },
          { name: '', bankNbfi: '', designation: '', sharePct: '', numShares: '' },
          { name: '', bankNbfi: '', designation: '', sharePct: '', numShares: '' },
        ];

  const finRows =
    bp?.section11Rows && bp.section11Rows.length > 0
      ? bp.section11Rows
      : [
          { year: '', grossSales: bp?.section11FinancialsNote || '', grossProfit: '', netProfit: '', totalAsset: '', totalLiabilities: '', netWorth: '' },
          { year: '', grossSales: '', grossProfit: '', netProfit: '', totalAsset: '', totalLiabilities: '', netWorth: '' },
        ];

  // Snapshot initializers when editing structured sections cleanly
  const currentSec2 = (cur: AiSectionsData['sec2LiabilityOurBank']): Section2AiData =>
    cur ?? {
      limit: ph['#56#'] || bp?.sec2Limit || '',
      validity: ph['#58#'] || bp?.sec2Validity || '',
      remarks: bp?.sec2Remarks || '',
      nonFundedNote: memo.sec2NonFundedNote || '',
      summary: memo.sectionSummaries.sec2LiabilityOurBank || '',
    };

  const currentSec3 = (cur: AiSectionsData['sec3LiabilityOtherUnits']): Section3AiData =>
    cur ?? {
      rows: rawSec3Rows.map((r) => ({ ...r })),
      summary: memo.sectionSummaries.sec3LiabilityOtherUnits || '',
    };

  const currentSec4 = (cur: AiSectionsData['sec4CibStatus']): Section4AiData =>
    cur ?? {
      cibDate: ph['#62#'] || bp?.section4CibDate || '',
      rows: sec4Rows.map((r) => ({ ...r })),
      summary: memo.sectionSummaries.sec4CibStatus || '',
    };

  const currentSec5 = (cur: AiSectionsData['sec5LegalAction']): Section5AiData =>
    cur ?? {
      statusNote: bp?.section5StatusNote || '',
      rows: sec5Rows.map((r) => ({ ...r })),
      summary: memo.sectionSummaries.sec5LegalAction || '',
    };

  const currentSec6 = (cur: AiSectionsData['sec6AuditCompliance']): Section6AiData =>
    cur ?? {
      statusNote: bp?.section6StatusNote || '',
      rows: sec6Rows.map((r) => ({ ...r })),
      summary: memo.sectionSummaries.sec6AuditCompliance || '',
    };

  const currentSec7 = (cur: AiSectionsData['sec7Security']): Section7AiData =>
    cur ?? {
      primaryNote: bp?.section7PrimaryNote || '',
      primaryRows: sec7Primary.map((r) => ({ ...r })),
      primaryTotal: bp?.section7PrimaryTotal || '',
      collateralNote: bp?.section7CollateralNote || '',
      collateralSchedules: [...sec7Schedules],
      valuationRows: sec7ValRows.map((r) => ({ ...r })),
      coveragePct: bp?.section7CoveragePct || '',
      othersNote: memo.sec7OthersNote || '',
      summary: memo.sectionSummaries.sec7Security || '',
    };

  const currentSec9 = (cur: AiSectionsData['sec9HoSanctions']): Section9AiData =>
    cur ?? {
      rows: sec9Rows.map((r) => ({ ...r })),
      summary: memo.sectionSummaries.sec9HoSanctions || '',
    };

  const currentSec10 = (cur: AiSectionsData['sec10ClientParticulars']): Section10AiData =>
    cur ?? {
      particulars: emptyParticulars(),
      directors: dirRows.map((d) => ({ ...d })),
      bankDirectors: bankDirRows.map((b) => ({ ...b })),
      summary: memo.sectionSummaries.sec10ClientParticulars || '',
    };

  const currentSec11 = (cur: AiSectionsData['sec11Financials']): Section11AiData =>
    cur ?? {
      financialsNote: bp?.section11FinancialsNote || '',
      rows: finRows.map((r) => ({ ...r })),
      summary: memo.sectionSummaries.sec11Financials || '',
    };

  const updateNarrativeAi = (
    key:
      | 'justification'
      | 'sec1NatureOfBusiness'
      | 'clientPrayer'
      | 'branchVisit'
      | 'sec13BranchRecommendation'
      | 'sec14IccdRecommendation'
      | 'samdObservation',
    val: string
  ) => {
    onUpdateAiSections((prev) => ({ ...prev, [key]: val }), key);
  };

  return (
    <div className="max-w-6xl mx-auto space-y-6">
      {/* Top Preview Controls */}
      <div className="bg-white border border-slate-200 rounded-lg p-4 flex flex-wrap items-center justify-between gap-4">
        <div>
          <h2 className="font-display text-lg font-bold text-slate-900">IC Memo Live Preview</h2>
          <p className="text-xs text-slate-500">
            Replicates the official Investment Committee A4 layout. Populate individual sections or generate all applicable sections together from the uploaded Branch Proposal.
          </p>
        </div>

        <div className="flex items-center gap-2.5 flex-wrap">
          <button
            type="button"
            onClick={() => void onScrutinizeAll()}
            disabled={!hasBranchProposal || batchScrutiny.running || scrutinizingSection !== null}
            title={
              !hasBranchProposal
                ? 'Upload the Branch Proposal on Manual Inputs first'
                : 'Scrutinize the uploaded Branch Proposal and populate all applicable sections in parallel (preserves manual edits)'
            }
            className="inline-flex items-center gap-2 px-4 py-2.5 text-sm font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-lg transition-colors cursor-pointer disabled:opacity-55 whitespace-nowrap shrink-0"
          >
            {batchScrutiny.running ? (
              <Loader2 className="w-4 h-4 animate-spin" />
            ) : (
              <Sparkles className="w-4 h-4 text-emerald-400" />
            )}
            <span>
              {batchScrutiny.running
                ? `Generating (${batchScrutiny.completedSections.length}/${batchScrutiny.total})...`
                : 'Generate and Fill Up All Together'}
            </span>
          </button>

          <button
            type="button"
            onClick={onApplyAllChanges}
            className={`inline-flex items-center gap-2 px-4 py-2.5 text-sm font-semibold rounded-lg border transition-colors cursor-pointer whitespace-nowrap shrink-0 ${
              hasPendingChanges
                ? 'text-amber-950 bg-amber-100 hover:bg-amber-200 border-amber-300'
                : 'text-slate-700 bg-slate-100 hover:bg-slate-200 border-slate-300'
            }`}
          >
            <RefreshCw className="w-4 h-4" />
            <span>Apply All Changes</span>
          </button>

          <button
            type="button"
            onClick={onDownloadDocx}
            disabled={downloading || downloadingPdf}
            className="inline-flex items-center gap-2 px-4 py-2.5 text-sm font-semibold text-white bg-emerald-800 hover:bg-emerald-900 rounded-lg transition-colors cursor-pointer disabled:opacity-60 whitespace-nowrap shrink-0"
          >
            {downloading ? (
              <Loader2 className="w-4 h-4 animate-spin" />
            ) : (
              <Download className="w-4 h-4" />
            )}
            <span>{downloading ? 'Building .docx...' : 'Download .docx'}</span>
          </button>

          <button
            type="button"
            onClick={onDownloadPdf}
            disabled={downloading || downloadingPdf}
            className="inline-flex items-center gap-2 px-4 py-2.5 text-sm font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-lg transition-colors cursor-pointer disabled:opacity-60 whitespace-nowrap shrink-0"
          >
            {downloadingPdf ? (
              <Loader2 className="w-4 h-4 animate-spin" />
            ) : (
              <Download className="w-4 h-4" />
            )}
            <span>{downloadingPdf ? 'Rendering .pdf...' : 'Download .pdf'}</span>
          </button>
        </div>
      </div>

      {/* Coordinated Batch Generation Progress & Result Summary (Status stays outside final memo document) */}
      {(batchScrutiny.running ||
        batchScrutiny.total > 0 ||
        batchScrutiny.failedSections.length > 0 ||
        batchScrutiny.insufficientSections.length > 0 ||
        batchScrutiny.skippedManualSections.length > 0) && (() => {
        const insufficientKeys = new Set(batchScrutiny.insufficientSections.map((i) => i.section));
        const failedKeys = new Set(batchScrutiny.failedSections.map((f) => f.section));
        const cleanlyPopulated = batchScrutiny.completedSections.filter(
          (s) => !insufficientKeys.has(s) && !failedKeys.has(s)
        );
        return (
          <div
            className={`border rounded-lg px-4 py-3 text-xs space-y-3 ${
              batchScrutiny.running
                ? 'bg-emerald-50/70 border-emerald-200 text-emerald-950'
                : batchScrutiny.failedSections.length > 0 ||
                    batchScrutiny.insufficientSections.length > 0
                  ? 'bg-amber-50/80 border-amber-200 text-slate-900'
                  : 'bg-emerald-50 border-emerald-200 text-emerald-950'
            }`}
          >
            <div className="flex items-center justify-between gap-3">
              <div className="flex items-center gap-2 font-semibold">
                {batchScrutiny.running ? (
                  <Loader2 className="w-4 h-4 animate-spin text-emerald-700 shrink-0" />
                ) : (
                  <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
                )}
                <span>
                  {batchScrutiny.running
                    ? `Generating and populating sections in parallel: ${
                        batchScrutiny.completedSections.length + batchScrutiny.failedSections.length
                      } of ${batchScrutiny.total} processed`
                    : batchScrutiny.total === 0
                      ? 'All sections already have manually edited content; nothing was overwritten.'
                      : `Batch generation finished: ${cleanlyPopulated.length} populated cleanly${
                          batchScrutiny.failedSections.length > 0
                            ? `, ${batchScrutiny.failedSections.length} require retry`
                            : ''
                        }${
                          batchScrutiny.insufficientSections.length > 0
                            ? `, ${batchScrutiny.insufficientSections.length} with unavailable source info`
                            : ''
                        }.`}
                </span>
              </div>

              {!batchScrutiny.running && (
                <button
                  type="button"
                  onClick={onDismissBatchSummary}
                  aria-label="Dismiss batch summary"
                  className="p-1 text-slate-500 hover:text-slate-800 rounded cursor-pointer"
                >
                  <X className="w-3.5 h-3.5" />
                </button>
              )}
            </div>

            {batchScrutiny.total > 0 && (
              <div className="w-full h-1.5 bg-white/80 border border-emerald-200/70 rounded-full overflow-hidden">
                <div
                  className="h-full bg-emerald-700 transition-all duration-200"
                  style={{
                    width: `${Math.min(
                      100,
                      Math.round(
                        ((batchScrutiny.completedSections.length +
                          batchScrutiny.failedSections.length) /
                          batchScrutiny.total) *
                          100
                      )
                    )}%`,
                  }}
                />
              </div>
            )}

            {cleanlyPopulated.length > 0 && (
              <div className="space-y-1">
                <p className="font-semibold text-[11px] text-emerald-900">
                  Successfully Populated Sections ({cleanlyPopulated.length}):
                </p>
                <div className="flex flex-wrap gap-1.5">
                  {cleanlyPopulated.map((s) => (
                    <span
                      key={s}
                      className="inline-flex items-center gap-1 px-2 py-0.5 bg-emerald-100/80 border border-emerald-300/70 text-emerald-950 rounded text-[11px]"
                    >
                      <CheckCircle2 className="w-3 h-3 text-emerald-700 shrink-0" />
                      <span>{SECTION_LABELS[s]}</span>
                    </span>
                  ))}
                </div>
              </div>
            )}

            {batchScrutiny.failedSections.length > 0 && (
              <div className="bg-white/90 border border-red-200 rounded-md p-2.5 space-y-1.5">
                <p className="font-semibold text-[11px] text-red-900">
                  Sections Requiring Retry ({batchScrutiny.failedSections.length}):
                </p>
                <div className="divide-y divide-red-100">
                  {batchScrutiny.failedSections.map((f) => {
                    const isRetrying = batchScrutiny.retryingSections.includes(f.section);
                    return (
                      <div
                        key={f.section}
                        className="py-1.5 first:pt-0 last:pb-0 flex flex-wrap items-center justify-between gap-2"
                      >
                        <div className="text-[11px] text-red-900">
                          <span className="font-semibold">{SECTION_LABELS[f.section]}:</span>{' '}
                          <span>{f.error}</span>
                        </div>
                        <button
                          type="button"
                          onClick={() => void onRetrySection(f.section)}
                          disabled={isRetrying || batchScrutiny.running}
                          className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-semibold text-white bg-amber-700 hover:bg-amber-800 rounded cursor-pointer disabled:opacity-60 whitespace-nowrap shrink-0"
                        >
                          {isRetrying ? (
                            <Loader2 className="w-3 h-3 animate-spin" />
                          ) : (
                            <RotateCcw className="w-3 h-3" />
                          )}
                          <span>{isRetrying ? 'Retrying...' : 'Retry'}</span>
                        </button>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {batchScrutiny.insufficientSections.length > 0 && (
              <div className="bg-white/90 border border-amber-300 rounded-md p-2.5 space-y-1.5">
                <p className="font-semibold text-[11px] text-amber-950">
                  Unavailable / Insufficient Source Information in Branch Proposal ({batchScrutiny.insufficientSections.length}):
                </p>
                <div className="divide-y divide-amber-100">
                  {batchScrutiny.insufficientSections.map((item) => {
                    const isRetrying = batchScrutiny.retryingSections.includes(item.section);
                    return (
                      <div
                        key={item.section}
                        className="py-1.5 first:pt-0 last:pb-0 flex flex-wrap items-center justify-between gap-2"
                      >
                        <div className="text-[11px] text-amber-950">
                          <span className="font-semibold">{SECTION_LABELS[item.section]}:</span>{' '}
                          <span>{item.reason}</span>
                        </div>
                        <button
                          type="button"
                          onClick={() => void onRetrySection(item.section)}
                          disabled={isRetrying || batchScrutiny.running}
                          className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-semibold text-amber-950 bg-amber-100 hover:bg-amber-200 border border-amber-300 rounded cursor-pointer disabled:opacity-60 whitespace-nowrap shrink-0"
                        >
                          {isRetrying ? (
                            <Loader2 className="w-3 h-3 animate-spin" />
                          ) : (
                            <RotateCcw className="w-3 h-3" />
                          )}
                          <span>{isRetrying ? 'Retrying...' : 'Retry'}</span>
                        </button>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {batchScrutiny.skippedManualSections.length > 0 && (
              <p className="text-[11px] text-slate-700">
                <span className="font-semibold">
                  Preserved manual edits ({batchScrutiny.skippedManualSections.length}):
                </span>{' '}
                {batchScrutiny.skippedManualSections.map((s) => SECTION_LABELS[s]).join(', ')}{' '}
                — left untouched without overwrite; use each section&apos;s individual AI button if you want to regenerate it.
              </p>
            )}
          </div>
        );
      })()}

      {!hasBranchProposal && (
        <div className="bg-amber-50 border border-amber-200 text-amber-900 rounded-lg px-4 py-2.5 text-xs flex items-center gap-2">
          <AlertTriangle className="w-4 h-4 text-amber-700 shrink-0" />
          <span>
            Upload the <span className="font-semibold">Branch Proposal</span> on the Manual Inputs tab before clicking <span className="font-semibold">“Scrutinize and Fill Up Through AI”</span> or <span className="font-semibold">“Generate and Fill Up All Together”</span>.
          </span>
        </div>
      )}

      {/* Default Format Template Control for IC Memo */}
      <TemplateManagerPanel
        templates={templates}
        onUploadTemplate={onUploadTemplate}
        onRemoveTemplate={onRemoveTemplate}
        singleDocType="ic"
      />

      {aiError && (
        <div className="bg-red-50 border border-red-200 text-red-800 rounded-lg px-4 py-3 text-sm">
          {aiError}
        </div>
      )}

      {/* Validation Panel */}
      <ValidationPanel
        missingFields={memo.missingFields}
        overrides={overrides}
        onOverrideChange={onOverrideChange}
      />

      {/* Official Document Paper Sheet (A4 595.32pt proportions) */}
      <div className="bg-slate-200/70 p-4 sm:p-8 rounded-lg overflow-x-auto">
        <div className="ic-memo-sheet max-w-[794px] mx-auto bg-white shadow-sm border border-slate-300 px-14 py-8 space-y-3 text-black">
          {/* Header: Bismillah + Logo + Bank Title + Exit with Waiver Box */}
          <div className="-mx-6">
            <div className="flex flex-col items-center">
              <img
                src="/assets/bismillah.jpg"
                alt="Bismillahir Rahmanir Rahim"
                referrerPolicy="no-referrer"
                className="h-5 object-contain mb-1"
              />
            </div>

            <div className="flex items-start justify-between gap-3">
              <img
                src="/assets/sib_logo.jpg"
                alt="Sammilito Islami Bank PLC"
                referrerPolicy="no-referrer"
                className="w-14 h-auto object-contain shrink-0"
              />
              <div className="text-center flex-1 leading-tight">
                <p className="text-[14pt] underline">SAMMILITO ISLAMI BANK PLC.</p>
                <p className="underline">
                  <span className="text-[9pt]">(</span>
                  <span className="text-[13pt]">State-owned Shariah-based Bank</span>
                  <span className="text-[9pt]">)</span>
                </p>
                <p className="text-[11pt] underline">UNION BANK UNIT</p>
                <p className="text-[9pt] underline">
                  Special Asset Management Division Head Office, Dhaka
                </p>
              </div>
              <div className="border border-black px-3 py-1 text-[12pt] font-bold whitespace-nowrap shrink-0 mt-1">
                Exit with Waiver
              </div>
            </div>
          </div>

          {/* Metadata Table */}
          <table className="w-[45%] table-fixed text-[10pt] border border-black">
            <colgroup>
              <col className="w-[28%]" />
              <col className="w-[7%]" />
              <col className="w-[65%]" />
            </colgroup>
            <tbody>
              <tr>
                <td className="border border-black px-2 py-0.5 font-bold">Dated</td>
                <td className="border border-black px-1.5 py-0.5 font-bold text-center">:</td>
                <td className="border border-black px-2 py-0.5 font-bold">
                  <Token ph={ph} code="#1#" />
                </td>
              </tr>
              <tr>
                <td className="border border-black px-2 py-0.5 font-bold">Memo no</td>
                <td className="border border-black px-1.5 py-0.5 font-bold text-center">:</td>
                <td className="border border-black px-2 py-0.5"></td>
              </tr>
              <tr>
                <td className="border border-black px-2 py-0.5 font-bold">To</td>
                <td className="border border-black px-1.5 py-0.5 font-bold text-center">:</td>
                <td className="border border-black px-2 py-0.5 font-bold">Investment Committee</td>
              </tr>
              <tr>
                <td className="border border-black px-2 py-0.5 font-bold">From</td>
                <td className="border border-black px-1.5 py-0.5 font-bold text-center">:</td>
                <td className="border border-black px-2 py-0.5 font-bold">SAMD</td>
              </tr>
            </tbody>
          </table>

          {/* A/C & Branch + Subject */}
          <table className="w-full table-fixed text-[10pt] border border-dotted border-black">
            <colgroup>
              <col className="w-[58%]" />
              <col className="w-[42%]" />
            </colgroup>
            <tbody>
              <tr>
                <td className="border border-dotted border-black px-2 py-0.5 font-bold text-center">
                  A/C: <Token ph={ph} code="#2#" /> (CIS: <Token ph={ph} code="#69#" />)
                </td>
                <td className="border border-dotted border-black px-2 py-0.5 font-bold text-center">
                  Branch: <Token ph={ph} code="#3#" />
                </td>
              </tr>
              <tr>
                <td colSpan={2} className="border border-dotted border-black px-2 py-1 font-bold leading-snug">
                  Subject: The customer proposed for full and final settlement of their liability by
                  allowing waiver of Tk.<Token ph={ph} code="#4#" /> along with profit/compensation to
                  be accrued up to the date of full and final settlement as follows:
                </td>
              </tr>
            </tbody>
          </table>

          <p className="text-right text-[10pt] font-bold">
            [All Figures in this Memo are in Million BDT Unless Stated Otherwise]
          </p>

          {/* Main Settlement Table */}
          <table className="w-full table-fixed text-[10pt] border border-dotted border-black">
            <colgroup>
              <col className="w-[12%]" />
              <col className="w-[8%]" />
              <col className="w-[7%]" />
              <col className="w-[21%]" />
              <col className="w-[11%]" />
              <col className="w-[11%]" />
              <col className="w-[11%]" />
              <col className="w-[19%]" />
            </colgroup>
            <thead>
              <tr>
                <th rowSpan={2} className="border border-dotted border-black px-1.5 py-0.5 text-center">
                  Mode
                </th>
                <th colSpan={2} className="border border-dotted border-black px-1.5 py-0.5 text-center">
                  Limit
                </th>
                <th rowSpan={2} className="border border-dotted border-black px-1.5 py-0.5 text-center">
                  Particulars
                </th>
                <th className="border border-dotted border-black px-1.5 py-0.5 text-center">
                  O/S as on <Token ph={ph} code="#68#" />
                </th>
                <th className="border border-dotted border-black px-1.5 py-0.5 text-center">
                  Proposed Waiver
                </th>
                <th className="border border-dotted border-black px-1.5 py-0.5 text-center">
                  Settlement Amount
                </th>
                <th rowSpan={2} className="border border-dotted border-black px-1.5 py-0.5 text-center">
                  Remarks
                </th>
              </tr>
              <tr>
                <th className="border border-dotted border-black px-1 py-0.5 text-center">Original</th>
                <th className="border border-dotted border-black px-1 py-0.5 text-center">RSD</th>
                <th className="border border-dotted border-black px-1 py-0.5 text-center">01</th>
                <th className="border border-dotted border-black px-1 py-0.5 text-center">02</th>
                <th className="border border-dotted border-black px-1 py-0.5 text-center">03 (1-2)</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td rowSpan={7} className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
                  <Token ph={ph} code="#5#" />
                </td>
                <td rowSpan={7} className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
                  <Token ph={ph} code="#6#" />
                </td>
                <td rowSpan={7} className="border border-dotted border-black px-1.5 py-0.5 text-center">
                  {bp?.limitRSD || '_'}
                </td>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-[9.5pt]">Principal</td>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
                  <Token ph={ph} code="#7#" />
                </td>
                <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                <td rowSpan={7} className="border border-dotted border-black px-1.5 py-0.5 font-bold leading-snug">
                  <div>1)Settlement Deadline:</div>
                  <div className="mb-2">
                    <Token ph={ph} code="#28#" />
                  </div>
                  <div>2) Break up</div>
                  <div>
                    Principal: <Token ph={ph} code="#7#" />
                  </div>
                  <div>
                    Charges: <Token ph={ph} code="#8#" />
                  </div>
                  <div>
                    COF: <Token ph={ph} code="#31#" />
                  </div>
                  <div className="mb-2">
                    Total: <Token ph={ph} code="#32#" />
                  </div>
                  <div>3) Security Value:</div>
                  <div>
                    MV: <Token ph={ph} code="#33#" />
                  </div>
                  <div>
                    FSV: <Token ph={ph} code="#34#" />
                  </div>
                </td>
              </tr>
              <tr>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-[9.5pt]">
                  Charges (Legal &amp; Others)
                </td>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
                  <Token ph={ph} code="#8#" />
                </td>
                <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                <td className="border border-dotted border-black px-1.5 py-0.5"></td>
              </tr>
              <tr>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-[9.5pt]">Gestation Rent</td>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center">-</td>
                <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                <td className="border border-dotted border-black px-1.5 py-0.5"></td>
              </tr>
              <tr>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-[9.5pt]">Charged Profit</td>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
                  <Token ph={ph} code="#10#" />
                </td>
                <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                <td className="border border-dotted border-black px-1.5 py-0.5"></td>
              </tr>
              <tr>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-[9.5pt]">Compensation</td>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
                  <Token ph={ph} code="#11#" />
                </td>
                <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                <td className="border border-dotted border-black px-1.5 py-0.5"></td>
              </tr>
              <tr>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-[9.5pt]">
                  Uncharged Compensation
                </td>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
                  <Token ph={ph} code="#12#" />
                </td>
                <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                <td className="border border-dotted border-black px-1.5 py-0.5"></td>
              </tr>
              <tr>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-[9.5pt]">
                  Uncharged Profit (for T/L)
                </td>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center">-</td>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center">-</td>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center">-</td>
              </tr>
              <tr>
                <td colSpan={4} className="border border-dotted border-black px-2 py-0.5 font-bold text-center">
                  Grand-Total
                </td>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
                  <Token ph={ph} code="#25#" />
                </td>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
                  <Token ph={ph} code="#4#" />
                </td>
                <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
                  <Token ph={ph} code="#27#" />
                </td>
                <td className="border border-dotted border-black px-1.5 py-0.5"></td>
              </tr>
            </tbody>
          </table>

          {/* Note */}
          <div className="border border-dotted border-black px-2 py-0.5 text-[10pt] font-bold">
            Note: Information Regarding Down payment (deposited Tk.<Token ph={ph} code="#35#" />{' '}
            Million in A/C: <Token ph={ph} code="#36#" /> on <Token ph={ph} code="#68#" />.)
          </div>

          {/* Key Points */}
          <div className="space-y-1">
            <p className="font-bold underline text-[10pt]">Key Points:</p>
            <p className="text-[10pt]">
              This Memo is placed in conformity with{' '}
              <span className="font-bold">BRPD 1 Circular Letter No. 23 dated 29.06.2026</span> under
              the following contexts:
            </p>
            <table className="w-full table-fixed text-[10pt] border border-dotted border-black">
              <colgroup>
                <col className="w-[5%]" />
                <col className="w-[42%]" />
                <col className="w-[11%]" />
                <col className="w-[5%]" />
                <col className="w-[22%]" />
                <col className="w-[15%]" />
              </colgroup>
              <thead>
                <tr>
                  <th colSpan={3} className="border border-dotted border-black px-2 py-0.5 text-left">
                    Particulars
                  </th>
                  <th colSpan={3} className="border border-dotted border-black px-2 py-0.5 text-left">
                    Waiver recommended by:
                  </th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">a.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">
                    Total [Profit+ Compensation+ Uncharged Comp.&amp; Profit]
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    : <Token ph={ph} code="#42#" />
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">d.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Client</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    : <Token ph={ph} code="#45#" />
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">b.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Cost of Fund</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    : <Token ph={ph} code="#31#" />
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">e.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Branch</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    : <Token ph={ph} code="#45#" />
                  </td>
                </tr>
                <tr>
                  <td rowSpan={4} className="border border-dotted border-black px-1.5 py-0.5">c.</td>
                  <td rowSpan={4} className="border border-dotted border-black px-1.5 py-0.5">
                    Maximum Waiver Considering COF [a-b]
                  </td>
                  <td rowSpan={4} className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    : <Token ph={ph} code="#44#" />
                  </td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">f.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">SAMD, HO</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    : <Token ph={ph} code="#47#" />
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">g.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">HOIC</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    : <Token ph={ph} code="#47#" />
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">h.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">MEC</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    : <Token ph={ph} code="#47#" />
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">i.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">IC &amp; CD</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    : <Token ph={ph} code="#47#" />
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* SECTION 0: Justification of The Proposed Facility */}
          <div className="space-y-1 text-[10pt]">
            <div className="flex flex-wrap items-start justify-between gap-2">
              <p className="font-bold underline pt-0.5">Justification of The Proposed Facility:</p>
              <SectionActionToolbar {...toolbarProps('justification')} />
            </div>

            <div className="pl-2 space-y-0.5">
              {isEditing('justification') ? (
                <textarea
                  rows={4}
                  value={aiSections.justification ?? memo.justificationLines}
                  onChange={(e) => updateNarrativeAi('justification', e.target.value)}
                  placeholder="Enter justification points or click Scrutinize and Fill Up Through AI..."
                  className="w-full font-memo text-[10pt] p-2 border border-emerald-600 rounded bg-emerald-50/20 focus:outline-none"
                />
              ) : (
                <p className="whitespace-pre-line">
                  a. <Token ph={ph} code="#51#" bold={false} />
                </p>
              )}
              <p>b. Investment was not created against any fraud and forgery.</p>
              <p>c. The client is not a willful defaulter.</p>
            </div>
          </div>

          <p className="font-bold text-[10pt]">
            Previous Board Decision(s) Regarding Rescheduling/Waiver/Written-Off: Nil
          </p>

          {/* SECTION 1: Nature of Business */}
          <div className="space-y-1">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div className="border border-dotted border-black px-2 py-0.5 font-bold text-[10pt] flex-1 min-w-[220px]">
                1. Nature of Business: <Token ph={ph} code="#54#" />
              </div>
              <SectionActionToolbar {...toolbarProps('sec1NatureOfBusiness')} />
            </div>
            {isEditing('sec1NatureOfBusiness') && (
              <div className="font-sans">
                <input
                  type="text"
                  value={aiSections.sec1NatureOfBusiness ?? ph['#54#'] ?? ''}
                  onChange={(e) => updateNarrativeAi('sec1NatureOfBusiness', e.target.value)}
                  placeholder="Enter Nature of Business..."
                  className="w-full text-xs px-2.5 py-1.5 border border-emerald-600 rounded bg-emerald-50/20 focus:outline-none"
                />
              </div>
            )}
          </div>

          {/* SECTION 2: Liability Position of the Clients and Its Sister Concerns with Our Bank */}
          <div className="space-y-1">
            <div className="flex flex-wrap items-start justify-between gap-2">
              <p className="font-bold text-[10pt] pt-0.5">
                2. Liability Position of the Clients and Its Sister Concerns with Our Bank as on{' '}
                <Token ph={ph} code="#68#" />:
              </p>
              <SectionActionToolbar {...toolbarProps('sec2LiabilityOurBank')} />
            </div>

            {isEditing('sec2LiabilityOurBank') ? (
              <div className="font-sans border border-emerald-600 rounded bg-emerald-50/20 p-2.5 space-y-2 text-xs">
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                  <label className="block">
                    <span className="text-[11px] font-medium text-slate-700">Limit (#56#)</span>
                    <input
                      type="text"
                      value={currentSec2(aiSections.sec2LiabilityOurBank).limit}
                      onChange={(e) =>
                        onUpdateAiSections(
                          (prev) => ({
                            ...prev,
                            sec2LiabilityOurBank: {
                              ...currentSec2(prev.sec2LiabilityOurBank),
                              limit: e.target.value,
                            },
                          }),
                          'sec2LiabilityOurBank'
                        )
                      }
                      className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                    />
                  </label>
                  <label className="block">
                    <span className="text-[11px] font-medium text-slate-700">Validity (#58#)</span>
                    <input
                      type="text"
                      value={currentSec2(aiSections.sec2LiabilityOurBank).validity}
                      onChange={(e) =>
                        onUpdateAiSections(
                          (prev) => ({
                            ...prev,
                            sec2LiabilityOurBank: {
                              ...currentSec2(prev.sec2LiabilityOurBank),
                              validity: e.target.value,
                            },
                          }),
                          'sec2LiabilityOurBank'
                        )
                      }
                      className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                    />
                  </label>
                  <label className="block">
                    <span className="text-[11px] font-medium text-slate-700">Funded Remarks</span>
                    <input
                      type="text"
                      value={currentSec2(aiSections.sec2LiabilityOurBank).remarks}
                      onChange={(e) =>
                        onUpdateAiSections(
                          (prev) => ({
                            ...prev,
                            sec2LiabilityOurBank: {
                              ...currentSec2(prev.sec2LiabilityOurBank),
                              remarks: e.target.value,
                            },
                          }),
                          'sec2LiabilityOurBank'
                        )
                      }
                      className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                    />
                  </label>
                </div>
                <label className="block">
                  <span className="text-[11px] font-medium text-slate-700">Non-Funded / Sister Concern Note</span>
                  <input
                    type="text"
                    value={currentSec2(aiSections.sec2LiabilityOurBank).nonFundedNote}
                    onChange={(e) =>
                      onUpdateAiSections(
                        (prev) => ({
                          ...prev,
                          sec2LiabilityOurBank: {
                            ...currentSec2(prev.sec2LiabilityOurBank),
                            nonFundedNote: e.target.value,
                          },
                        }),
                        'sec2LiabilityOurBank'
                      )
                    }
                    className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                  />
                </label>
                <label className="block">
                  <span className="text-[11px] font-medium text-slate-700">Section 2 AI/Manual Summary Note</span>
                  <textarea
                    rows={2}
                    value={currentSec2(aiSections.sec2LiabilityOurBank).summary}
                    onChange={(e) =>
                      onUpdateAiSections(
                        (prev) => ({
                          ...prev,
                          sec2LiabilityOurBank: {
                            ...currentSec2(prev.sec2LiabilityOurBank),
                            summary: e.target.value,
                          },
                        }),
                        'sec2LiabilityOurBank'
                      )
                    }
                    className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                  />
                </label>
              </div>
            ) : (
              <>
                <table className="w-full table-fixed text-[10pt] border border-dotted border-black">
                  <colgroup>
                    <col className="w-[14%]" />
                    <col className="w-[17%]" />
                    <col className="w-[10%]" />
                    <col className="w-[11%]" />
                    <col className="w-[11%]" />
                    <col className="w-[10%]" />
                    <col className="w-[17%]" />
                    <col className="w-[10%]" />
                  </colgroup>
                  <thead>
                    <tr>
                      <th className="border border-dotted border-black px-1 py-0.5 text-left">Name of the Concern</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-left">Mode of Investment</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Limit</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Outstanding</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Validity</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Overdue</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">
                        Recovery in existing deals till date [Mode wise]
                      </th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-left">Remarks</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td rowSpan={4} className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                        <Token ph={ph} code="#2#" />
                      </td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                        <div>
                          <Token ph={ph} code="#55#" />
                        </div>
                        <div>Funded</div>
                      </td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
                        <Token ph={ph} code="#56#" />
                      </td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
                        <Token ph={ph} code="#25#" />
                      </td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
                        <Token ph={ph} code="#58#" />
                      </td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
                        <Token ph={ph} code="#59#" />
                      </td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
                        <Token ph={ph} code="#60#" />
                      </td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">
                        {bp?.sec2Remarks || ''}
                      </td>
                    </tr>
                    <tr className="h-4">
                      <td className="border border-dotted border-black px-1.5 py-0.5">&nbsp;</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                      <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                      <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                      <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                      <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                      <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                    </tr>
                    <tr className="h-4">
                      <td className="border border-dotted border-black px-1.5 py-0.5">&nbsp;</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                      <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                      <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                      <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                      <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                      <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                    </tr>
                    <tr>
                      <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">Non Funded</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                      <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                      <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                      <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                      <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">
                        {memo.sec2NonFundedNote || ''}
                      </td>
                    </tr>
                    <tr>
                      <td colSpan={2} className="border border-dotted border-black px-1.5 py-0.5 font-bold text-center">
                        Total
                      </td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
                        <Token ph={ph} code="#56#" />
                      </td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
                        <Token ph={ph} code="#25#" />
                      </td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
                        <Token ph={ph} code="#58#" />
                      </td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
                        <Token ph={ph} code="#59#" />
                      </td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center font-bold">
                        <Token ph={ph} code="#60#" />
                      </td>
                      <td className="border border-dotted border-black px-1.5 py-0.5"></td>
                    </tr>
                  </tbody>
                </table>
              </>
            )}
          </div>

          <div className="border-t border-dashed border-slate-300 pt-2 text-right font-sans text-[10px] text-slate-400">
            Page Break — Section 3 begins on Page 2
          </div>

          {/* SECTION 3: Liability with Other Units of SIB */}
          <div className="space-y-1">
            <div className="flex flex-wrap items-start justify-between gap-2">
              <p className="font-bold underline text-[10pt] pt-0.5">
                3. Liability of the Client &amp; Its Sister/Allied Concerns with Other Units of SIB:
              </p>
              <SectionActionToolbar {...toolbarProps('sec3LiabilityOtherUnits')} />
            </div>

            {isEditing('sec3LiabilityOtherUnits') ? (
              <div className="font-sans border border-emerald-600 rounded bg-emerald-50/20 p-2 space-y-2 text-xs">
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse">
                    <thead>
                      <tr className="text-[10px] text-slate-700">
                        <th className="p-1 text-left">SL</th>
                        <th className="p-1 text-left">Concern</th>
                        <th className="p-1 text-left">Banks/NBFIs</th>
                        <th className="p-1 text-left">Facility</th>
                        <th className="p-1 text-left">Limit</th>
                        <th className="p-1 text-left">Outstanding</th>
                        <th className="p-1 text-left">Overdue/EOL</th>
                        <th className="p-1 text-left">Expiry</th>
                        <th className="p-1 text-left">CL Status</th>
                        <th className="p-1"></th>
                      </tr>
                    </thead>
                    <tbody>
                      {currentSec3(aiSections.sec3LiabilityOtherUnits).rows.map((r, i) => (
                        <tr key={i}>
                          {(
                            [
                              'sl',
                              'nameOfConcern',
                              'bankOrNbfi',
                              'natureOfFacility',
                              'limit',
                              'outstanding',
                              'overdueEol',
                              'expiryDate',
                              'clStatus',
                            ] as const
                          ).map((field) => (
                            <td key={field} className="p-0.5">
                              <input
                                type="text"
                                value={r[field]}
                                onChange={(e) =>
                                  onUpdateAiSections((prev) => {
                                    const base = currentSec3(prev.sec3LiabilityOtherUnits);
                                    const nextRows = base.rows.map((row, idx) =>
                                      idx === i ? { ...row, [field]: e.target.value } : row
                                    );
                                    return {
                                      ...prev,
                                      sec3LiabilityOtherUnits: { ...base, rows: nextRows },
                                    };
                                  }, 'sec3LiabilityOtherUnits')
                                }
                                className="w-full px-1 py-0.5 border border-slate-300 rounded bg-white text-[11px]"
                              />
                            </td>
                          ))}
                          <td className="p-0.5">
                            <button
                              type="button"
                              onClick={() =>
                                onUpdateAiSections((prev) => {
                                  const base = currentSec3(prev.sec3LiabilityOtherUnits);
                                  return {
                                    ...prev,
                                    sec3LiabilityOtherUnits: {
                                      ...base,
                                      rows: base.rows.filter((_, idx) => idx !== i),
                                    },
                                  };
                                }, 'sec3LiabilityOtherUnits')
                              }
                              className="p-1 text-red-600 hover:bg-red-50 rounded cursor-pointer"
                            >
                              <Trash2 className="w-3 h-3" />
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <button
                  type="button"
                  onClick={() =>
                    onUpdateAiSections((prev) => {
                      const base = currentSec3(prev.sec3LiabilityOtherUnits);
                      return {
                        ...prev,
                        sec3LiabilityOtherUnits: {
                          ...base,
                          rows: [
                            ...base.rows,
                            {
                              sl: `${String.fromCharCode(97 + base.rows.length)}.`,
                              nameOfConcern: '',
                              bankOrNbfi: '',
                              natureOfFacility: '',
                              limit: '',
                              outstanding: '',
                              overdueEol: '',
                              expiryDate: '',
                              clStatus: '',
                            },
                          ],
                        },
                      };
                    }, 'sec3LiabilityOtherUnits')
                  }
                  className="inline-flex items-center gap-1 px-2 py-1 text-[11px] font-medium text-emerald-900 bg-white border border-emerald-400 rounded cursor-pointer"
                >
                  <Plus className="w-3 h-3" /> Add Row
                </button>
                <label className="block">
                  <span className="text-[11px] font-medium text-slate-700">Section 3 Summary / Note</span>
                  <textarea
                    rows={2}
                    value={currentSec3(aiSections.sec3LiabilityOtherUnits).summary}
                    onChange={(e) =>
                      onUpdateAiSections(
                        (prev) => ({
                          ...prev,
                          sec3LiabilityOtherUnits: {
                            ...currentSec3(prev.sec3LiabilityOtherUnits),
                            summary: e.target.value,
                          },
                        }),
                        'sec3LiabilityOtherUnits'
                      )
                    }
                    className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                  />
                </label>
              </div>
            ) : (
              <>
                <table className="w-full table-fixed text-[9.5pt] border border-dotted border-black">
                  <colgroup>
                    <col className="w-[5%]" />
                    <col className="w-[13%]" />
                    <col className="w-[17%]" />
                    <col className="w-[11%]" />
                    <col className="w-[8%]" />
                    <col className="w-[17%]" />
                    <col className="w-[11%]" />
                    <col className="w-[10%]" />
                    <col className="w-[8%]" />
                  </colgroup>
                  <thead>
                    <tr>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">SL</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-left">Name of Concern</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-left">Name of the Banks/NBFIs</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-left">Nature of Facility</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Limit</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Outstanding with profit/Interest</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Overdue/EOL</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Expiry Date</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">CL Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {rawSec3Rows.map((r, i) => (
                      <tr key={i}>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.sl}</td>
                        <td className="border border-dotted border-black px-1 py-0.5">{r.nameOfConcern}</td>
                        <td className="border border-dotted border-black px-1 py-0.5">{r.bankOrNbfi}</td>
                        <td className="border border-dotted border-black px-1 py-0.5">{r.natureOfFacility}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.limit}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.outstanding}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.overdueEol}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.expiryDate}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.clStatus}</td>
                      </tr>
                    ))}
                    {!hasSec3SubTotal && (
                      <tr>
                        <td colSpan={4} className="border border-dotted border-black px-1 py-0.5 font-bold text-center">
                          Sub-Total
                        </td>
                        <td className="border border-dotted border-black px-1 py-0.5"></td>
                        <td className="border border-dotted border-black px-1 py-0.5"></td>
                        <td className="border border-dotted border-black px-1 py-0.5"></td>
                        <td className="border border-dotted border-black px-1 py-0.5"></td>
                        <td className="border border-dotted border-black px-1 py-0.5"></td>
                      </tr>
                    )}
                    {!hasSec3GrandTotal && (
                      <tr>
                        <td colSpan={4} className="border border-dotted border-black px-1 py-0.5 font-bold text-center">
                          Grand Total
                        </td>
                        <td className="border border-dotted border-black px-1 py-0.5"></td>
                        <td className="border border-dotted border-black px-1 py-0.5"></td>
                        <td className="border border-dotted border-black px-1 py-0.5"></td>
                        <td className="border border-dotted border-black px-1 py-0.5"></td>
                        <td className="border border-dotted border-black px-1 py-0.5"></td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </>
            )}
          </div>

          {/* SECTION 4: CIB Status */}
          <div className="space-y-1">
            <div className="flex flex-wrap items-start justify-between gap-2">
              <p className="font-bold text-[10pt] pt-0.5">
                <span className="underline">4. CIB Status:</span> Report dated{' '}
                <Token ph={ph} code="#62#" /> shows:
              </p>
              <SectionActionToolbar {...toolbarProps('sec4CibStatus')} />
            </div>

            {isEditing('sec4CibStatus') ? (
              <div className="font-sans border border-emerald-600 rounded bg-emerald-50/20 p-2 space-y-2 text-xs">
                <label className="block max-w-xs">
                  <span className="text-[11px] font-medium text-slate-700">CIB Date (Branch Proposal Sec. 4)</span>
                  <input
                    type="text"
                    value={currentSec4(aiSections.sec4CibStatus).cibDate}
                    onChange={(e) =>
                      onUpdateAiSections(
                        (prev) => ({
                          ...prev,
                          sec4CibStatus: { ...currentSec4(prev.sec4CibStatus), cibDate: e.target.value },
                        }),
                        'sec4CibStatus'
                      )
                    }
                    className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                  />
                </label>
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse">
                    <thead>
                      <tr className="text-[10px] text-slate-700">
                        <th className="p-1 text-left">STATUS</th>
                        <th className="p-1 text-left">Unit Borrower</th>
                        <th className="p-1 text-left">Unit Guarantor</th>
                        <th className="p-1 text-left">Other Borrower</th>
                        <th className="p-1 text-left">Other Guarantor</th>
                        <th className="p-1 text-left">Remarks</th>
                      </tr>
                    </thead>
                    <tbody>
                      {currentSec4(aiSections.sec4CibStatus).rows.map((r, i) => (
                        <tr key={i}>
                          <td className="p-0.5 font-medium text-[11px]">{r.status}</td>
                          {(
                            ['unitBorrower', 'unitGuarantor', 'otherBorrower', 'otherGuarantor', 'remarks'] as const
                          ).map((field) => (
                            <td key={field} className="p-0.5">
                              <input
                                type="text"
                                value={r[field]}
                                onChange={(e) =>
                                  onUpdateAiSections((prev) => {
                                    const base = currentSec4(prev.sec4CibStatus);
                                    return {
                                      ...prev,
                                      sec4CibStatus: {
                                        ...base,
                                        rows: base.rows.map((row, idx) =>
                                          idx === i ? { ...row, [field]: e.target.value } : row
                                        ),
                                      },
                                    };
                                  }, 'sec4CibStatus')
                                }
                                className="w-full px-1 py-0.5 border border-slate-300 rounded bg-white text-[11px]"
                              />
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <label className="block">
                  <span className="text-[11px] font-medium text-slate-700">Section 4 Summary / Note</span>
                  <textarea
                    rows={2}
                    value={currentSec4(aiSections.sec4CibStatus).summary}
                    onChange={(e) =>
                      onUpdateAiSections(
                        (prev) => ({
                          ...prev,
                          sec4CibStatus: { ...currentSec4(prev.sec4CibStatus), summary: e.target.value },
                        }),
                        'sec4CibStatus'
                      )
                    }
                    className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                  />
                </label>
              </div>
            ) : (
              <>
                <table className="w-full table-fixed text-[10pt] border border-dotted border-black">
                  <colgroup>
                    <col className="w-[18%]" />
                    <col className="w-[13%]" />
                    <col className="w-[13%]" />
                    <col className="w-[13%]" />
                    <col className="w-[13%]" />
                    <col className="w-[30%]" />
                  </colgroup>
                  <thead>
                    <tr>
                      <th rowSpan={2} className="border border-dotted border-black px-1 py-0.5 text-left">STATUS</th>
                      <th colSpan={2} className="border border-dotted border-black px-1 py-0.5 text-center">
                        With respective unit
                      </th>
                      <th colSpan={2} className="border border-dotted border-black px-1 py-0.5 text-center">
                        With Other Bank
                      </th>
                      <th rowSpan={2} className="border border-dotted border-black px-1 py-0.5 text-center">
                        Remarks (Pls. give concern wise Breakup)
                      </th>
                    </tr>
                    <tr>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Borrower</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Guarantor</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Borrower</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Guarantor</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sec4Rows.map((r, i) => (
                      <tr key={i} className={/TOTAL/i.test(r.status) ? 'font-bold' : ''}>
                        <td className="border border-dotted border-black px-1.5 py-0.5">{r.status}</td>
                        <td className="border border-dotted border-black px-1.5 py-0.5 text-center">{r.unitBorrower}</td>
                        <td className="border border-dotted border-black px-1.5 py-0.5 text-center">{r.unitGuarantor}</td>
                        <td className="border border-dotted border-black px-1.5 py-0.5 text-center">{r.otherBorrower}</td>
                        <td className="border border-dotted border-black px-1.5 py-0.5 text-center">{r.otherGuarantor}</td>
                        <td className="border border-dotted border-black px-1.5 py-0.5">{r.remarks}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </>
            )}
          </div>

          {/* SECTION 5: Legal Action Status */}
          <div className="space-y-1">
            <div className="flex flex-wrap items-start justify-between gap-2">
              <p className="font-bold text-[10pt] pt-0.5">
                <span className="underline">5. Legal Action Status as on Memo Date:</span>{' '}
                {bp?.section5StatusNote || ''}
              </p>
              <SectionActionToolbar {...toolbarProps('sec5LegalAction')} />
            </div>

            {isEditing('sec5LegalAction') ? (
              <div className="font-sans border border-emerald-600 rounded bg-emerald-50/20 p-2 space-y-2 text-xs">
                <label className="block max-w-xs">
                  <span className="text-[11px] font-medium text-slate-700">Header Status Note (e.g. Nil, N/A)</span>
                  <input
                    type="text"
                    value={currentSec5(aiSections.sec5LegalAction).statusNote}
                    onChange={(e) =>
                      onUpdateAiSections(
                        (prev) => ({
                          ...prev,
                          sec5LegalAction: { ...currentSec5(prev.sec5LegalAction), statusNote: e.target.value },
                        }),
                        'sec5LegalAction'
                      )
                    }
                    className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                  />
                </label>
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse">
                    <thead>
                      <tr className="text-[10px] text-slate-700">
                        <th className="p-1 text-left">Sl</th>
                        <th className="p-1 text-left">Type of Case</th>
                        <th className="p-1 text-left">Case No.</th>
                        <th className="p-1 text-left">Filing date</th>
                        <th className="p-1 text-left">Amount</th>
                        <th className="p-1 text-left">Prev date</th>
                        <th className="p-1 text-left">Next date</th>
                        <th className="p-1 text-left">Present status</th>
                        <th className="p-1"></th>
                      </tr>
                    </thead>
                    <tbody>
                      {currentSec5(aiSections.sec5LegalAction).rows.map((r, i) => (
                        <tr key={i}>
                          {(
                            [
                              'sl',
                              'typeOfCase',
                              'caseNo',
                              'filingDate',
                              'amount',
                              'previousDate',
                              'nextDate',
                              'presentStatus',
                            ] as const
                          ).map((field) => (
                            <td key={field} className="p-0.5">
                              <input
                                type="text"
                                value={r[field]}
                                onChange={(e) =>
                                  onUpdateAiSections((prev) => {
                                    const base = currentSec5(prev.sec5LegalAction);
                                    return {
                                      ...prev,
                                      sec5LegalAction: {
                                        ...base,
                                        rows: base.rows.map((row, idx) =>
                                          idx === i ? { ...row, [field]: e.target.value } : row
                                        ),
                                      },
                                    };
                                  }, 'sec5LegalAction')
                                }
                                className="w-full px-1 py-0.5 border border-slate-300 rounded bg-white text-[11px]"
                              />
                            </td>
                          ))}
                          <td className="p-0.5">
                            <button
                              type="button"
                              onClick={() =>
                                onUpdateAiSections((prev) => {
                                  const base = currentSec5(prev.sec5LegalAction);
                                  return {
                                    ...prev,
                                    sec5LegalAction: {
                                      ...base,
                                      rows: base.rows.filter((_, idx) => idx !== i),
                                    },
                                  };
                                }, 'sec5LegalAction')
                              }
                              className="p-1 text-red-600 hover:bg-red-50 rounded cursor-pointer"
                            >
                              <Trash2 className="w-3 h-3" />
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <button
                  type="button"
                  onClick={() =>
                    onUpdateAiSections((prev) => {
                      const base = currentSec5(prev.sec5LegalAction);
                      return {
                        ...prev,
                        sec5LegalAction: {
                          ...base,
                          rows: [
                            ...base.rows,
                            {
                              sl: String(base.rows.length + 1),
                              typeOfCase: '',
                              caseNo: '',
                              filingDate: '',
                              amount: '',
                              previousDate: '',
                              nextDate: '',
                              presentStatus: '',
                            },
                          ],
                        },
                      };
                    }, 'sec5LegalAction')
                  }
                  className="inline-flex items-center gap-1 px-2 py-1 text-[11px] font-medium text-emerald-900 bg-white border border-emerald-400 rounded cursor-pointer"
                >
                  <Plus className="w-3 h-3" /> Add Case Row
                </button>
                <label className="block">
                  <span className="text-[11px] font-medium text-slate-700">Section 5 Summary / Note</span>
                  <textarea
                    rows={2}
                    value={currentSec5(aiSections.sec5LegalAction).summary}
                    onChange={(e) =>
                      onUpdateAiSections(
                        (prev) => ({
                          ...prev,
                          sec5LegalAction: { ...currentSec5(prev.sec5LegalAction), summary: e.target.value },
                        }),
                        'sec5LegalAction'
                      )
                    }
                    className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                  />
                </label>
              </div>
            ) : (
              <>
                <table className="w-full table-fixed text-[9.5pt] border border-dotted border-black">
                  <colgroup>
                    <col className="w-[5%]" />
                    <col className="w-[16%]" />
                    <col className="w-[13%]" />
                    <col className="w-[12%]" />
                    <col className="w-[12%]" />
                    <col className="w-[12%]" />
                    <col className="w-[12%]" />
                    <col className="w-[18%]" />
                  </colgroup>
                  <thead>
                    <tr>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Sl</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-left">Type of Case</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-left">Case No.</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Filling date</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Amount</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Previous date</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Next date</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-left">Present status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sec5Rows.map((r, i) => {
                      const isEmpty = !r.sl && !r.typeOfCase && !r.caseNo;
                      return (
                        <tr key={i} className={isEmpty ? 'h-9' : ''}>
                          <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.sl}</td>
                          <td className="border border-dotted border-black px-1 py-0.5">{r.typeOfCase}</td>
                          <td className="border border-dotted border-black px-1 py-0.5">{r.caseNo}</td>
                          <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.filingDate}</td>
                          <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.amount}</td>
                          <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.previousDate}</td>
                          <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.nextDate}</td>
                          <td className="border border-dotted border-black px-1 py-0.5">{r.presentStatus}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </>
            )}
          </div>

          {/* SECTION 6: Audit, Compliance and Regulatory Authority Related Status */}
          <div className="space-y-1">
            <div className="flex flex-wrap items-start justify-between gap-2">
              <p className="font-bold text-[10pt] pt-0.5">
                <span className="underline">
                  6. Audit, Compliance and Regulatory Authority Related Status:
                </span>{' '}
                {bp?.section6StatusNote || ''}
              </p>
              <SectionActionToolbar {...toolbarProps('sec6AuditCompliance')} />
            </div>

            {isEditing('sec6AuditCompliance') ? (
              <div className="font-sans border border-emerald-600 rounded bg-emerald-50/20 p-2 space-y-2 text-xs">
                <label className="block max-w-xs">
                  <span className="text-[11px] font-medium text-slate-700">Header Status Note</span>
                  <input
                    type="text"
                    value={currentSec6(aiSections.sec6AuditCompliance).statusNote}
                    onChange={(e) =>
                      onUpdateAiSections(
                        (prev) => ({
                          ...prev,
                          sec6AuditCompliance: {
                            ...currentSec6(prev.sec6AuditCompliance),
                            statusNote: e.target.value,
                          },
                        }),
                        'sec6AuditCompliance'
                      )
                    }
                    className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                  />
                </label>
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse">
                    <thead>
                      <tr className="text-[10px] text-slate-700">
                        <th className="p-1 text-left">SL</th>
                        <th className="p-1 text-left">Authority</th>
                        <th className="p-1 text-left">Last Inspection Date</th>
                        <th className="p-1 text-left">Objection (Yes/No)</th>
                        <th className="p-1 text-left">Objections</th>
                        <th className="p-1 text-left">Branch Reply</th>
                      </tr>
                    </thead>
                    <tbody>
                      {currentSec6(aiSections.sec6AuditCompliance).rows.map((r, i) => (
                        <tr key={i}>
                          <td className="p-0.5 text-[11px]">{r.sl}</td>
                          <td className="p-0.5 text-[11px]">{r.authority}</td>
                          {(
                            ['lastInspectionDate', 'objectionStatus', 'objections', 'branchReply'] as const
                          ).map((field) => (
                            <td key={field} className="p-0.5">
                              <input
                                type="text"
                                value={r[field]}
                                onChange={(e) =>
                                  onUpdateAiSections((prev) => {
                                    const base = currentSec6(prev.sec6AuditCompliance);
                                    return {
                                      ...prev,
                                      sec6AuditCompliance: {
                                        ...base,
                                        rows: base.rows.map((row, idx) =>
                                          idx === i ? { ...row, [field]: e.target.value } : row
                                        ),
                                      },
                                    };
                                  }, 'sec6AuditCompliance')
                                }
                                className="w-full px-1 py-0.5 border border-slate-300 rounded bg-white text-[11px]"
                              />
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <label className="block">
                  <span className="text-[11px] font-medium text-slate-700">Section 6 Summary / Note</span>
                  <textarea
                    rows={2}
                    value={currentSec6(aiSections.sec6AuditCompliance).summary}
                    onChange={(e) =>
                      onUpdateAiSections(
                        (prev) => ({
                          ...prev,
                          sec6AuditCompliance: {
                            ...currentSec6(prev.sec6AuditCompliance),
                            summary: e.target.value,
                          },
                        }),
                        'sec6AuditCompliance'
                      )
                    }
                    className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                  />
                </label>
              </div>
            ) : (
              <>
                <table className="w-full table-fixed text-[9.5pt] border border-dotted border-black">
                  <colgroup>
                    <col className="w-[5%]" />
                    <col className="w-[19%]" />
                    <col className="w-[16%]" />
                    <col className="w-[14%]" />
                    <col className="w-[23%]" />
                    <col className="w-[23%]" />
                  </colgroup>
                  <thead>
                    <tr>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">SL</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-left">Name of the Authority</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Last Inspection / Report Date</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Objection Status (Yes/No)</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-left">Objections (If yes)</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-left">Branch Reply/ Current Position</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sec6Rows.map((r, i) => (
                      <tr key={i}>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.sl}</td>
                        <td className="border border-dotted border-black px-1 py-0.5">{r.authority}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.lastInspectionDate}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.objectionStatus}</td>
                        <td className="border border-dotted border-black px-1 py-0.5">{r.objections}</td>
                        <td className="border border-dotted border-black px-1 py-0.5">{r.branchReply}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </>
            )}
          </div>

          {/* SECTION 7: Security */}
          <div className="space-y-1.5 text-[10pt]">
            <div className="flex flex-wrap items-start justify-between gap-2">
              <p className="font-bold underline text-[10pt] pt-0.5">7. Security:</p>
              <SectionActionToolbar {...toolbarProps('sec7Security')} />
            </div>

            {isEditing('sec7Security') ? (
              <div className="font-sans border border-emerald-600 rounded bg-emerald-50/20 p-2.5 space-y-2.5 text-xs">
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                  <label className="block">
                    <span className="text-[11px] font-medium text-slate-700">a. Primary Note</span>
                    <input
                      type="text"
                      value={currentSec7(aiSections.sec7Security).primaryNote}
                      onChange={(e) =>
                        onUpdateAiSections(
                          (prev) => ({
                            ...prev,
                            sec7Security: { ...currentSec7(prev.sec7Security), primaryNote: e.target.value },
                          }),
                          'sec7Security'
                        )
                      }
                      className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                    />
                  </label>
                  <label className="block">
                    <span className="text-[11px] font-medium text-slate-700">Primary Security Total</span>
                    <input
                      type="text"
                      value={currentSec7(aiSections.sec7Security).primaryTotal}
                      onChange={(e) =>
                        onUpdateAiSections(
                          (prev) => ({
                            ...prev,
                            sec7Security: { ...currentSec7(prev.sec7Security), primaryTotal: e.target.value },
                          }),
                          'sec7Security'
                        )
                      }
                      className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                    />
                  </label>
                  <label className="block">
                    <span className="text-[11px] font-medium text-slate-700">b. Collateral Note</span>
                    <input
                      type="text"
                      value={currentSec7(aiSections.sec7Security).collateralNote}
                      onChange={(e) =>
                        onUpdateAiSections(
                          (prev) => ({
                            ...prev,
                            sec7Security: { ...currentSec7(prev.sec7Security), collateralNote: e.target.value },
                          }),
                          'sec7Security'
                        )
                      }
                      className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                    />
                  </label>
                </div>

                <div className="space-y-1">
                  <p className="text-[11px] font-semibold text-slate-700">Primary Security Particulars</p>
                  {currentSec7(aiSections.sec7Security).primaryRows.map((r, i) => (
                    <div key={i} className="grid grid-cols-12 gap-1.5">
                      <input
                        type="text"
                        value={r.particulars}
                        onChange={(e) =>
                          onUpdateAiSections((prev) => {
                            const base = currentSec7(prev.sec7Security);
                            return {
                              ...prev,
                              sec7Security: {
                                ...base,
                                primaryRows: base.primaryRows.map((row, idx) =>
                                  idx === i ? { ...row, particulars: e.target.value } : row
                                ),
                              },
                            };
                          }, 'sec7Security')
                        }
                        className="col-span-8 px-2 py-1 border border-slate-300 rounded bg-white text-[11px]"
                      />
                      <input
                        type="text"
                        value={r.value}
                        placeholder="Value"
                        onChange={(e) =>
                          onUpdateAiSections((prev) => {
                            const base = currentSec7(prev.sec7Security);
                            return {
                              ...prev,
                              sec7Security: {
                                ...base,
                                primaryRows: base.primaryRows.map((row, idx) =>
                                  idx === i ? { ...row, value: e.target.value } : row
                                ),
                              },
                            };
                          }, 'sec7Security')
                        }
                        className="col-span-4 px-2 py-1 border border-slate-300 rounded bg-white text-[11px]"
                      />
                    </div>
                  ))}
                </div>

                <label className="block">
                  <span className="text-[11px] font-medium text-slate-700">Collateral Schedules (one per line)</span>
                  <textarea
                    rows={3}
                    value={currentSec7(aiSections.sec7Security).collateralSchedules.join('\n')}
                    onChange={(e) =>
                      onUpdateAiSections(
                        (prev) => ({
                          ...prev,
                          sec7Security: {
                            ...currentSec7(prev.sec7Security),
                            collateralSchedules: e.target.value.split('\n'),
                          },
                        }),
                        'sec7Security'
                      )
                    }
                    className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                  />
                </label>

                <div className="overflow-x-auto">
                  <p className="text-[11px] font-semibold text-slate-700 mb-1">Collateral Valuation Table</p>
                  <table className="w-full border-collapse text-[10px]">
                    <thead>
                      <tr>
                        <th className="p-0.5 text-left">Schedule</th>
                        <th className="p-0.5">Orig Br MV</th>
                        <th className="p-0.5">Orig Br FSV</th>
                        <th className="p-0.5">Orig Surv MV</th>
                        <th className="p-0.5">Orig Surv FSV</th>
                        <th className="p-0.5">Pres Br MV</th>
                        <th className="p-0.5">Pres Br FSV</th>
                        <th className="p-0.5">Pres Surv MV</th>
                        <th className="p-0.5">Pres Surv FSV</th>
                        <th className="p-0.5">AV MV</th>
                        <th className="p-0.5">AV FSV</th>
                      </tr>
                    </thead>
                    <tbody>
                      {currentSec7(aiSections.sec7Security).valuationRows.map((r, i) => (
                        <tr key={i}>
                          {(
                            [
                              'schedule',
                              'origBranchMv',
                              'origBranchFsv',
                              'origSurvMv',
                              'origSurvFsv',
                              'presBranchMv',
                              'presBranchFsv',
                              'presSurvMv',
                              'presSurvFsv',
                              'presAvMv',
                              'presAvFsv',
                            ] as const
                          ).map((field) => (
                            <td key={field} className="p-0.5">
                              <input
                                type="text"
                                value={r[field]}
                                onChange={(e) =>
                                  onUpdateAiSections((prev) => {
                                    const base = currentSec7(prev.sec7Security);
                                    return {
                                      ...prev,
                                      sec7Security: {
                                        ...base,
                                        valuationRows: base.valuationRows.map((row, idx) =>
                                          idx === i ? { ...row, [field]: e.target.value } : row
                                        ),
                                      },
                                    };
                                  }, 'sec7Security')
                                }
                                className="w-full px-1 py-0.5 border border-slate-300 rounded bg-white text-[10px]"
                              />
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  <label className="block">
                    <span className="text-[11px] font-medium text-slate-700">Coverage (%)</span>
                    <input
                      type="text"
                      value={currentSec7(aiSections.sec7Security).coveragePct}
                      onChange={(e) =>
                        onUpdateAiSections(
                          (prev) => ({
                            ...prev,
                            sec7Security: { ...currentSec7(prev.sec7Security), coveragePct: e.target.value },
                          }),
                          'sec7Security'
                        )
                      }
                      className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                    />
                  </label>
                  <label className="block">
                    <span className="text-[11px] font-medium text-slate-700">c. Others (Additional Note)</span>
                    <input
                      type="text"
                      value={currentSec7(aiSections.sec7Security).othersNote}
                      onChange={(e) =>
                        onUpdateAiSections(
                          (prev) => ({
                            ...prev,
                            sec7Security: { ...currentSec7(prev.sec7Security), othersNote: e.target.value },
                          }),
                          'sec7Security'
                        )
                      }
                      className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                    />
                  </label>
                </div>

                <label className="block">
                  <span className="text-[11px] font-medium text-slate-700">Section 7 Summary / Assessment</span>
                  <textarea
                    rows={2}
                    value={currentSec7(aiSections.sec7Security).summary}
                    onChange={(e) =>
                      onUpdateAiSections(
                        (prev) => ({
                          ...prev,
                          sec7Security: { ...currentSec7(prev.sec7Security), summary: e.target.value },
                        }),
                        'sec7Security'
                      )
                    }
                    className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                  />
                </label>
              </div>
            ) : (
              <>
                <p className="font-bold pl-2">
                  <span className="underline">a. Primary:</span> {bp?.section7PrimaryNote || ''}
                </p>
                <table className="w-full table-fixed text-[10pt] border border-dotted border-black">
                  <colgroup>
                    <col className="w-[75%]" />
                    <col className="w-[25%]" />
                  </colgroup>
                  <thead>
                    <tr>
                      <th className="border border-dotted border-black px-2 py-0.5 text-left">Particulars</th>
                      <th className="border border-dotted border-black px-2 py-0.5 text-center">Value</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sec7Primary.map((r, i) => (
                      <tr key={i}>
                        <td className="border border-dotted border-black px-2 py-0.5">{r.particulars}</td>
                        <td className="border border-dotted border-black px-2 py-0.5 text-center">{r.value}</td>
                      </tr>
                    ))}
                    <tr className="font-bold">
                      <td className="border border-dotted border-black px-2 py-0.5 text-center">
                        Total value of primary security
                      </td>
                      <td className="border border-dotted border-black px-2 py-0.5 text-center">
                        {bp?.section7PrimaryTotal || ''}
                      </td>
                    </tr>
                  </tbody>
                </table>

                <p className="font-bold pl-2 pt-1">
                  <span className="underline">b. Collateral:</span> {bp?.section7CollateralNote || ''}
                </p>
                <div className="pl-2 space-y-0.5 text-[10pt]">
                  {sec7Schedules.map((s, i) => (
                    <p key={i}>{s}</p>
                  ))}
                </div>

                <table className="w-full table-fixed text-[9pt] border border-dotted border-black">
                  <colgroup>
                    <col className="w-[13%]" />
                    <col className="w-[8.7%]" />
                    <col className="w-[8.7%]" />
                    <col className="w-[8.7%]" />
                    <col className="w-[8.7%]" />
                    <col className="w-[8.7%]" />
                    <col className="w-[8.7%]" />
                    <col className="w-[8.7%]" />
                    <col className="w-[8.7%]" />
                    <col className="w-[8.7%]" />
                    <col className="w-[8.7%]" />
                  </colgroup>
                  <thead>
                    <tr>
                      <th rowSpan={4} className="border border-dotted border-black px-1 py-1 align-middle text-center">
                        Schedule
                      </th>
                      <th colSpan={10} className="border border-dotted border-black px-1 py-0.5 text-center">
                        Value of the mortgaged properties as assessed
                      </th>
                    </tr>
                    <tr>
                      <th colSpan={4} className="border border-dotted border-black px-1 py-0.5 text-center">
                        At the time of allowing investment
                      </th>
                      <th colSpan={6} className="border border-dotted border-black px-1 py-0.5 text-center">
                        At present
                      </th>
                    </tr>
                    <tr>
                      <th colSpan={2} className="border border-dotted border-black px-1 py-0.5 text-center">
                        By Branch<br />on………..
                      </th>
                      <th colSpan={2} className="border border-dotted border-black px-1 py-0.5 text-center">
                        Surveyor<br />on……………..
                      </th>
                      <th colSpan={2} className="border border-dotted border-black px-1 py-0.5 text-center">
                        By Branch<br />on…………….
                      </th>
                      <th colSpan={2} className="border border-dotted border-black px-1 py-0.5 text-center">
                        Surveyor<br />on………….
                      </th>
                      <th colSpan={2} className="border border-dotted border-black px-1 py-0.5 text-center">
                        By AV Cell<br />on……………
                      </th>
                    </tr>
                    <tr>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">MV</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">FSV</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">MV</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">FSV</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">MV</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">FSV</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">MV</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">FSV</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">MV</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">FSV</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sec7ValRows.map((r, i) => (
                      <tr key={i} className={/Grand/i.test(r.schedule) ? 'font-bold' : ''}>
                        <td className="border border-dotted border-black px-1 py-0.5">{r.schedule}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.origBranchMv}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.origBranchFsv}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.origSurvMv}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.origSurvFsv}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.presBranchMv}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.presBranchFsv}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.presSurvMv}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.presSurvFsv}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.presAvMv}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.presAvFsv}</td>
                      </tr>
                    ))}
                    <tr>
                      <td colSpan={11} className="border border-dotted border-black px-2 py-0.5 text-center font-bold">
                        Outstanding to Lowest Security value Coverage (%): {bp?.section7CoveragePct || '……………..%'}
                      </td>
                    </tr>
                  </tbody>
                </table>

                <div className="pl-2 font-bold text-[10pt]">
                  <p>c. Others:</p>
                  <p className="pl-5">a. Usual Charge documents.</p>
                  <p className="pl-5">b. Acceptance of Branch Sanction by the Client.</p>
                  {memo.sec7OthersNote && (
                    <p className="pl-5 font-normal whitespace-pre-line">{memo.sec7OthersNote}</p>
                  )}
                </div>
              </>
            )}
          </div>

          {/* SECTION 8: Source of Repayment */}
          <div className="space-y-1 text-[10pt]">
            <div className="flex flex-wrap items-start justify-between gap-2">
              <p className="font-bold underline pt-0.5">8. Source of Repayment:</p>
              <SectionActionToolbar {...toolbarProps('sec8SourceOfRepayment')} />
            </div>
            {isEditing('sec8SourceOfRepayment') ? (
              <textarea
                rows={3}
                value={sourceOfRepaymentManual || aiSections.sec8SourceOfRepayment || ''}
                onChange={(e) => onManualSourceOfRepaymentChange(e.target.value)}
                placeholder="Enter source of repayment or click Scrutinize and Fill Up Through AI..."
                className="w-full font-memo text-[10pt] p-2 border border-emerald-600 rounded bg-emerald-50/20 focus:outline-none"
              />
            ) : memo.sourceOfRepayment ? (
              <p className="whitespace-pre-line">{memo.sourceOfRepayment}</p>
            ) : (
              <p className="text-slate-400 italic">[Enter in Manual Inputs or click Scrutinize and Fill Up Through AI]</p>
            )}
          </div>

          {/* SECTION 9: Sequence of Head Office Sanction(S) */}
          <div className="space-y-1">
            <div className="flex flex-wrap items-start justify-between gap-2">
              <p className="font-bold underline text-[10pt] pt-0.5">9. Sequence of Head Office Sanction(S):</p>
              <SectionActionToolbar {...toolbarProps('sec9HoSanctions')} />
            </div>

            {isEditing('sec9HoSanctions') ? (
              <div className="font-sans border border-emerald-600 rounded bg-emerald-50/20 p-2 space-y-2 text-xs">
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse">
                    <thead>
                      <tr className="text-[10px] text-slate-700">
                        <th className="p-1 text-left">SL</th>
                        <th className="p-1 text-left">HO Sanction Reference No.</th>
                        <th className="p-1 text-left">Date</th>
                        <th className="p-1 text-left">Limit</th>
                        <th className="p-1 text-left">Remarks</th>
                        <th className="p-1"></th>
                      </tr>
                    </thead>
                    <tbody>
                      {currentSec9(aiSections.sec9HoSanctions).rows.map((r, i) => (
                        <tr key={i}>
                          {(['sl', 'referenceNo', 'date', 'limit', 'remarks'] as const).map((field) => (
                            <td key={field} className="p-0.5">
                              <input
                                type="text"
                                value={r[field]}
                                onChange={(e) =>
                                  onUpdateAiSections((prev) => {
                                    const base = currentSec9(prev.sec9HoSanctions);
                                    return {
                                      ...prev,
                                      sec9HoSanctions: {
                                        ...base,
                                        rows: base.rows.map((row, idx) =>
                                          idx === i ? { ...row, [field]: e.target.value } : row
                                        ),
                                      },
                                    };
                                  }, 'sec9HoSanctions')
                                }
                                className="w-full px-1 py-0.5 border border-slate-300 rounded bg-white text-[11px]"
                              />
                            </td>
                          ))}
                          <td className="p-0.5">
                            <button
                              type="button"
                              onClick={() =>
                                onUpdateAiSections((prev) => {
                                  const base = currentSec9(prev.sec9HoSanctions);
                                  return {
                                    ...prev,
                                    sec9HoSanctions: {
                                      ...base,
                                      rows: base.rows.filter((_, idx) => idx !== i),
                                    },
                                  };
                                }, 'sec9HoSanctions')
                              }
                              className="p-1 text-red-600 hover:bg-red-50 rounded cursor-pointer"
                            >
                              <Trash2 className="w-3 h-3" />
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <button
                  type="button"
                  onClick={() =>
                    onUpdateAiSections((prev) => {
                      const base = currentSec9(prev.sec9HoSanctions);
                      return {
                        ...prev,
                        sec9HoSanctions: {
                          ...base,
                          rows: [
                            ...base.rows,
                            {
                              sl: `${String.fromCharCode(97 + base.rows.length)}.`,
                              referenceNo: '',
                              date: '',
                              limit: '',
                              remarks: '',
                            },
                          ],
                        },
                      };
                    }, 'sec9HoSanctions')
                  }
                  className="inline-flex items-center gap-1 px-2 py-1 text-[11px] font-medium text-emerald-900 bg-white border border-emerald-400 rounded cursor-pointer"
                >
                  <Plus className="w-3 h-3" /> Add Sanction Row
                </button>
                <label className="block">
                  <span className="text-[11px] font-medium text-slate-700">Section 9 Summary / Note</span>
                  <textarea
                    rows={2}
                    value={currentSec9(aiSections.sec9HoSanctions).summary}
                    onChange={(e) =>
                      onUpdateAiSections(
                        (prev) => ({
                          ...prev,
                          sec9HoSanctions: { ...currentSec9(prev.sec9HoSanctions), summary: e.target.value },
                        }),
                        'sec9HoSanctions'
                      )
                    }
                    className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                  />
                </label>
              </div>
            ) : (
              <>
                <table className="w-full table-fixed text-[10pt] border border-dotted border-black">
                  <colgroup>
                    <col className="w-[7%]" />
                    <col className="w-[36%]" />
                    <col className="w-[16%]" />
                    <col className="w-[15%]" />
                    <col className="w-[26%]" />
                  </colgroup>
                  <thead>
                    <tr>
                      <th className="border border-dotted border-black px-1.5 py-0.5 text-center">SL</th>
                      <th className="border border-dotted border-black px-1.5 py-0.5 text-left">HO Sanction Reference No.</th>
                      <th className="border border-dotted border-black px-1.5 py-0.5 text-center">Date</th>
                      <th className="border border-dotted border-black px-1.5 py-0.5 text-center">Limit</th>
                      <th className="border border-dotted border-black px-1.5 py-0.5 text-left">Remarks</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sec9Rows.map((r, i) => (
                      <tr key={i}>
                        <td className="border border-dotted border-black px-1.5 py-0.5 text-center">{r.sl}</td>
                        <td className="border border-dotted border-black px-1.5 py-0.5">{r.referenceNo}</td>
                        <td className="border border-dotted border-black px-1.5 py-0.5 text-center">{r.date}</td>
                        <td className="border border-dotted border-black px-1.5 py-0.5 text-center">{r.limit}</td>
                        <td className="border border-dotted border-black px-1.5 py-0.5">{r.remarks}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </>
            )}
          </div>

          {/* SECTION 10: Brief Particulars of The Client */}
          <div className="space-y-2">
            <div className="flex flex-wrap items-start justify-between gap-2">
              <p className="font-bold underline text-[10pt] pt-0.5">10. Brief Particulars of The Client:</p>
              <SectionActionToolbar {...toolbarProps('sec10ClientParticulars')} />
            </div>

            {isEditing('sec10ClientParticulars') ? (
              <div className="font-sans border border-emerald-600 rounded bg-emerald-50/20 p-2.5 space-y-3 text-xs">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  {(
                    [
                      ['nameOfClient', 'a. Name of the Client'],
                      ['uniqueCustomerId', 'b. Unique Customer ID'],
                      ['awcdMsdAccount', 'c. AWCD/MSD Account with opening date'],
                      ['businessTypeNature', 'd. Business Type/Nature'],
                      ['businessManagementType', 'e. Business Management Type'],
                      ['clientType', 'f. Client Type'],
                      ['keyPersonPhone', 'g. Key Person with Phone No.'],
                      ['address', 'g. Address'],
                      ['officeBusinessAddress', 'h. Office/Business Address'],
                      ['factoryAddress', 'h. Factory Address'],
                      ['purposeOfInvestment', 'i. Purpose of Investment'],
                      ['limitOfInvestment', 'j. Limit of Investment'],
                      ['modeOfInvestment', 'k. Mode of Investment'],
                      ['lastVisitDate', 'l. Last Visit Date'],
                      ['stockPosition', 'm. Stock Position'],
                      ['fixedAsset', 'n. Fixed Asset'],
                      ['businessNetWorth', 'o. Business Net Worth'],
                      ['insuranceCoverage', 'p. Insurance Coverage'],
                      ['tradeLicenseValidity', 'q. Trade License & Validity'],
                      ['otherRegulatoryDocuments', 'r. Other Regulatory Documents'],
                      ['dateOfOriginalInvestment', 's. Date of Original Investment'],
                      ['firstDisbursementManager', 't. 1st Disbursement Manager'],
                      ['firstDateOfClassification', 'u. First Date of Classification with Status'],
                    ] as Array<[keyof Section10Particulars, string]>
                  ).map(([field, label]) => (
                    <label key={field} className="block">
                      <span className="text-[11px] font-medium text-slate-700">{label}</span>
                      <input
                        type="text"
                        value={currentSec10(aiSections.sec10ClientParticulars).particulars[field]}
                        onChange={(e) =>
                          onUpdateAiSections((prev) => {
                            const base = currentSec10(prev.sec10ClientParticulars);
                            return {
                              ...prev,
                              sec10ClientParticulars: {
                                ...base,
                                particulars: {
                                  ...base.particulars,
                                  [field]: e.target.value,
                                },
                              },
                            };
                          }, 'sec10ClientParticulars')
                        }
                        className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white text-[11px]"
                      />
                    </label>
                  ))}
                </div>

                <div>
                  <p className="text-[11px] font-semibold text-slate-700 mb-1">
                    Proprietor / Partner / Directors
                  </p>
                  <div className="overflow-x-auto">
                    <table className="w-full border-collapse">
                      <thead>
                        <tr className="text-[10px] text-slate-700">
                          <th className="p-1 text-left">Name</th>
                          <th className="p-1 text-left">Age</th>
                          <th className="p-1 text-left">Education</th>
                          <th className="p-1 text-left">Designation</th>
                          <th className="p-1 text-left">Share (%)</th>
                          <th className="p-1 text-left">Experience</th>
                          <th className="p-1 text-left">Net worth</th>
                          <th className="p-1"></th>
                        </tr>
                      </thead>
                      <tbody>
                        {currentSec10(aiSections.sec10ClientParticulars).directors.map((d, i) => (
                          <tr key={i}>
                            {(
                              ['name', 'age', 'education', 'designation', 'sharePct', 'experience', 'netWorth'] as const
                            ).map((field) => (
                              <td key={field} className="p-0.5">
                                <input
                                  type="text"
                                  value={d[field]}
                                  onChange={(e) =>
                                    onUpdateAiSections((prev) => {
                                      const base = currentSec10(prev.sec10ClientParticulars);
                                      return {
                                        ...prev,
                                        sec10ClientParticulars: {
                                          ...base,
                                          directors: base.directors.map((row, idx) =>
                                            idx === i ? { ...row, [field]: e.target.value } : row
                                          ),
                                        },
                                      };
                                    }, 'sec10ClientParticulars')
                                  }
                                  className="w-full px-1 py-0.5 border border-slate-300 rounded bg-white text-[11px]"
                                />
                              </td>
                            ))}
                            <td className="p-0.5">
                              <button
                                type="button"
                                onClick={() =>
                                  onUpdateAiSections((prev) => {
                                    const base = currentSec10(prev.sec10ClientParticulars);
                                    return {
                                      ...prev,
                                      sec10ClientParticulars: {
                                        ...base,
                                        directors: base.directors.filter((_, idx) => idx !== i),
                                      },
                                    };
                                  }, 'sec10ClientParticulars')
                                }
                                className="p-1 text-red-600 hover:bg-red-50 rounded cursor-pointer"
                              >
                                <Trash2 className="w-3 h-3" />
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <button
                    type="button"
                    onClick={() =>
                      onUpdateAiSections((prev) => {
                        const base = currentSec10(prev.sec10ClientParticulars);
                        return {
                          ...prev,
                          sec10ClientParticulars: {
                            ...base,
                            directors: [
                              ...base.directors,
                              {
                                name: '',
                                age: '',
                                education: '',
                                designation: '',
                                sharePct: '',
                                experience: '',
                                netWorth: '',
                              },
                            ],
                          },
                        };
                      }, 'sec10ClientParticulars')
                    }
                    className="mt-1 inline-flex items-center gap-1 px-2 py-0.5 text-[11px] font-medium text-emerald-900 bg-white border border-emerald-400 rounded cursor-pointer"
                  >
                    <Plus className="w-3 h-3" /> Add Director
                  </button>
                </div>

                <div>
                  <p className="text-[11px] font-semibold text-slate-700 mb-1">
                    Director of any Bank / FI / Insurance Company
                  </p>
                  <div className="overflow-x-auto">
                    <table className="w-full border-collapse">
                      <thead>
                        <tr className="text-[10px] text-slate-700">
                          <th className="p-1 text-left">Name</th>
                          <th className="p-1 text-left">Bank/NBFIs</th>
                          <th className="p-1 text-left">Designation</th>
                          <th className="p-1 text-left">Share (%)</th>
                          <th className="p-1 text-left">No. of Shares</th>
                          <th className="p-1"></th>
                        </tr>
                      </thead>
                      <tbody>
                        {currentSec10(aiSections.sec10ClientParticulars).bankDirectors.map((d, i) => (
                          <tr key={i}>
                            {(['name', 'bankNbfi', 'designation', 'sharePct', 'numShares'] as const).map((field) => (
                              <td key={field} className="p-0.5">
                                <input
                                  type="text"
                                  value={d[field]}
                                  onChange={(e) =>
                                    onUpdateAiSections((prev) => {
                                      const base = currentSec10(prev.sec10ClientParticulars);
                                      return {
                                        ...prev,
                                        sec10ClientParticulars: {
                                          ...base,
                                          bankDirectors: base.bankDirectors.map((row, idx) =>
                                            idx === i ? { ...row, [field]: e.target.value } : row
                                          ),
                                        },
                                      };
                                    }, 'sec10ClientParticulars')
                                  }
                                  className="w-full px-1 py-0.5 border border-slate-300 rounded bg-white text-[11px]"
                                />
                              </td>
                            ))}
                            <td className="p-0.5">
                              <button
                                type="button"
                                onClick={() =>
                                  onUpdateAiSections((prev) => {
                                    const base = currentSec10(prev.sec10ClientParticulars);
                                    return {
                                      ...prev,
                                      sec10ClientParticulars: {
                                        ...base,
                                        bankDirectors: base.bankDirectors.filter((_, idx) => idx !== i),
                                      },
                                    };
                                  }, 'sec10ClientParticulars')
                                }
                                className="p-1 text-red-600 hover:bg-red-50 rounded cursor-pointer"
                              >
                                <Trash2 className="w-3 h-3" />
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <button
                    type="button"
                    onClick={() =>
                      onUpdateAiSections((prev) => {
                        const base = currentSec10(prev.sec10ClientParticulars);
                        return {
                          ...prev,
                          sec10ClientParticulars: {
                            ...base,
                            bankDirectors: [
                              ...base.bankDirectors,
                              { name: '', bankNbfi: '', designation: '', sharePct: '', numShares: '' },
                            ],
                          },
                        };
                      }, 'sec10ClientParticulars')
                    }
                    className="mt-1 inline-flex items-center gap-1 px-2 py-0.5 text-[11px] font-medium text-emerald-900 bg-white border border-emerald-400 rounded cursor-pointer"
                  >
                    <Plus className="w-3 h-3" /> Add Bank Director
                  </button>
                </div>

                <label className="block">
                  <span className="text-[11px] font-medium text-slate-700">Section 10 Summary / Note</span>
                  <textarea
                    rows={2}
                    value={currentSec10(aiSections.sec10ClientParticulars).summary}
                    onChange={(e) =>
                      onUpdateAiSections(
                        (prev) => ({
                          ...prev,
                          sec10ClientParticulars: {
                            ...currentSec10(prev.sec10ClientParticulars),
                            summary: e.target.value,
                          },
                        }),
                        'sec10ClientParticulars'
                      )
                    }
                    className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                  />
                </label>
              </div>
            ) : (
              <>
                <table className="w-full table-fixed text-[10pt] border border-dotted border-black">
                  <colgroup>
                    <col className="w-[6%]" />
                    <col className="w-[37%]" />
                    <col className="w-[4%]" />
                    <col className="w-[53%]" />
                  </colgroup>
                  <tbody>
                    <tr>
                      <td className="border border-dotted border-black px-1.5 py-0.5">a.</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">Name of the Client</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center">:</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">{sp?.nameOfClient || ''}</td>
                    </tr>
                    <tr>
                      <td className="border border-dotted border-black px-1.5 py-0.5">b.</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">Unique Customer ID</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center">:</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">{sp?.uniqueCustomerId || ''}</td>
                    </tr>
                    <tr>
                      <td className="border border-dotted border-black px-1.5 py-0.5">c.</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">AWCD/MSD Account with opening date</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center">:</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">{sp?.awcdMsdAccount || ''}</td>
                    </tr>
                    <tr>
                      <td className="border border-dotted border-black px-1.5 py-0.5">d.</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">Business Type/Nature</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center">:</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">{sp?.businessTypeNature || ''}</td>
                    </tr>
                    <tr>
                      <td className="border border-dotted border-black px-1.5 py-0.5">e.</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">Business Management Type</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center">:</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">{sp?.businessManagementType || ''}</td>
                    </tr>
                    <tr>
                      <td className="border border-dotted border-black px-1.5 py-0.5">f.</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">Client Type</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center">:</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">{sp?.clientType || ''}</td>
                    </tr>
                    <tr>
                      <td rowSpan={2} className="border border-dotted border-black px-1.5 py-0.5 align-middle">g.</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">Key Person with Phone No.</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center">:</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">{sp?.keyPersonPhone || ''}</td>
                    </tr>
                    <tr>
                      <td className="border border-dotted border-black px-1.5 py-0.5">Address</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center">:</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">{sp?.address || ''}</td>
                    </tr>
                    <tr>
                      <td rowSpan={2} className="border border-dotted border-black px-1.5 py-0.5 align-middle">h.</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">Office/Business Address</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center">:</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">{sp?.officeBusinessAddress || ''}</td>
                    </tr>
                    <tr>
                      <td className="border border-dotted border-black px-1.5 py-0.5">Factory Address</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center">:</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">{sp?.factoryAddress || ''}</td>
                    </tr>
                    {[
                      ['i.', 'Purpose of Investment', sp?.purposeOfInvestment || ''],
                      ['j.', 'Limit of Investment', sp?.limitOfInvestment || ''],
                      ['k.', 'Mode of Investment', sp?.modeOfInvestment || ''],
                      ['l.', 'Last Visit Date', sp?.lastVisitDate || ''],
                      ['m.', 'Stock Position', sp?.stockPosition || ''],
                      ['n.', 'Fixed Asset', sp?.fixedAsset || ''],
                      ['o.', 'Business Net Worth', sp?.businessNetWorth || ''],
                      ['p.', 'Insurance Coverage', sp?.insuranceCoverage || ''],
                      ['q.', 'Trade License & Validity', sp?.tradeLicenseValidity || ''],
                      ['r', 'Other Regulatory Documents', sp?.otherRegulatoryDocuments || ''],
                      ['s', 'Date of Original Investment', sp?.dateOfOriginalInvestment || ''],
                    ].map(([idx, label, val], i) => (
                      <tr key={i}>
                        <td className="border border-dotted border-black px-1.5 py-0.5">{idx}</td>
                        <td className="border border-dotted border-black px-1.5 py-0.5">{label}</td>
                        <td className="border border-dotted border-black px-1.5 py-0.5 text-center">:</td>
                        <td className="border border-dotted border-black px-1.5 py-0.5">{val}</td>
                      </tr>
                    ))}
                    <tr>
                      <td className="border border-dotted border-black px-1.5 py-0.5">t.</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">
                        1<sup>st</sup> Disbursement Manager
                      </td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center">:</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">{sp?.firstDisbursementManager || ''}</td>
                    </tr>
                    <tr>
                      <td className="border border-dotted border-black px-1.5 py-0.5">u.</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">First Date of Classification with Status</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5 text-center">:</td>
                      <td className="border border-dotted border-black px-1.5 py-0.5">{sp?.firstDateOfClassification || ''}</td>
                    </tr>
                  </tbody>
                </table>

                <p className="font-bold underline text-[10pt]">
                  • Particular of the Proprietor/Partner/Directors:
                </p>
                <table className="w-full table-fixed text-[9.5pt] border border-dotted border-black">
                  <colgroup>
                    <col className="w-[22%]" />
                    <col className="w-[8%]" />
                    <col className="w-[15%]" />
                    <col className="w-[16%]" />
                    <col className="w-[12%]" />
                    <col className="w-[15%]" />
                    <col className="w-[12%]" />
                  </colgroup>
                  <thead>
                    <tr>
                      <th className="border border-dotted border-black px-1 py-0.5 text-left">Name of the Directors</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Age</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-left">Education</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-left">Designation</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Share (%)</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Experience</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Net worth</th>
                    </tr>
                  </thead>
                  <tbody>
                    {dirRows.map((d, i) => (
                      <tr key={i} className={!d.name && !d.age ? 'h-6' : ''}>
                        <td className="border border-dotted border-black px-1 py-0.5">{d.name}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{d.age}</td>
                        <td className="border border-dotted border-black px-1 py-0.5">{d.education}</td>
                        <td className="border border-dotted border-black px-1 py-0.5">{d.designation}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{d.sharePct}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{d.experience}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{d.netWorth}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>

                <p className="font-bold underline text-[10pt]">
                  • Remarks: Director of any Bank/ FI/ Insurance company:
                </p>
                <table className="w-full table-fixed text-[9.5pt] border border-dotted border-black">
                  <colgroup>
                    <col className="w-[22%]" />
                    <col className="w-[22%]" />
                    <col className="w-[18%]" />
                    <col className="w-[19%]" />
                    <col className="w-[19%]" />
                  </colgroup>
                  <thead>
                    <tr>
                      <th className="border border-dotted border-black px-1 py-0.5 text-left">Name of the Director</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-left">Name of the bank/NBFIs</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-left">Designation/Position</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Percentage of Shares</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Number of Shares</th>
                    </tr>
                  </thead>
                  <tbody>
                    {bankDirRows.map((d, i) => (
                      <tr key={i} className={!d.name && !d.bankNbfi ? 'h-6' : ''}>
                        <td className="border border-dotted border-black px-1 py-0.5">{d.name}</td>
                        <td className="border border-dotted border-black px-1 py-0.5">{d.bankNbfi}</td>
                        <td className="border border-dotted border-black px-1 py-0.5">{d.designation}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{d.sharePct}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{d.numShares}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </>
            )}
          </div>

          {/* SECTION 11: Financials of The Client */}
          <div className="space-y-1">
            <div className="flex flex-wrap items-start justify-between gap-2">
              <p className="font-bold underline text-[10pt] pt-0.5">11. Financials of The Client:</p>
              <SectionActionToolbar {...toolbarProps('sec11Financials')} />
            </div>

            {isEditing('sec11Financials') ? (
              <div className="font-sans border border-emerald-600 rounded bg-emerald-50/20 p-2 space-y-2 text-xs">
                <label className="block max-w-xs">
                  <span className="text-[11px] font-medium text-slate-700">Financials Note (e.g. Nil, N/A)</span>
                  <input
                    type="text"
                    value={currentSec11(aiSections.sec11Financials).financialsNote}
                    onChange={(e) =>
                      onUpdateAiSections(
                        (prev) => ({
                          ...prev,
                          sec11Financials: {
                            ...currentSec11(prev.sec11Financials),
                            financialsNote: e.target.value,
                          },
                        }),
                        'sec11Financials'
                      )
                    }
                    className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                  />
                </label>
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse">
                    <thead>
                      <tr className="text-[10px] text-slate-700">
                        <th className="p-1 text-left">Year</th>
                        <th className="p-1 text-left">Gross Sales</th>
                        <th className="p-1 text-left">Gross Profit</th>
                        <th className="p-1 text-left">Net Profit</th>
                        <th className="p-1 text-left">Total Asset</th>
                        <th className="p-1 text-left">Total liabilities</th>
                        <th className="p-1 text-left">Net Worth</th>
                        <th className="p-1"></th>
                      </tr>
                    </thead>
                    <tbody>
                      {currentSec11(aiSections.sec11Financials).rows.map((r, i) => (
                        <tr key={i}>
                          {(
                            [
                              'year',
                              'grossSales',
                              'grossProfit',
                              'netProfit',
                              'totalAsset',
                              'totalLiabilities',
                              'netWorth',
                            ] as const
                          ).map((field) => (
                            <td key={field} className="p-0.5">
                              <input
                                type="text"
                                value={r[field]}
                                onChange={(e) =>
                                  onUpdateAiSections((prev) => {
                                    const base = currentSec11(prev.sec11Financials);
                                    return {
                                      ...prev,
                                      sec11Financials: {
                                        ...base,
                                        rows: base.rows.map((row, idx) =>
                                          idx === i ? { ...row, [field]: e.target.value } : row
                                        ),
                                      },
                                    };
                                  }, 'sec11Financials')
                                }
                                className="w-full px-1 py-0.5 border border-slate-300 rounded bg-white text-[11px]"
                              />
                            </td>
                          ))}
                          <td className="p-0.5">
                            <button
                              type="button"
                              onClick={() =>
                                onUpdateAiSections((prev) => {
                                  const base = currentSec11(prev.sec11Financials);
                                  return {
                                    ...prev,
                                    sec11Financials: {
                                      ...base,
                                      rows: base.rows.filter((_, idx) => idx !== i),
                                    },
                                  };
                                }, 'sec11Financials')
                              }
                              className="p-1 text-red-600 hover:bg-red-50 rounded cursor-pointer"
                            >
                              <Trash2 className="w-3 h-3" />
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <button
                  type="button"
                  onClick={() =>
                    onUpdateAiSections((prev) => {
                      const base = currentSec11(prev.sec11Financials);
                      return {
                        ...prev,
                        sec11Financials: {
                          ...base,
                          rows: [
                            ...base.rows,
                            {
                              year: '',
                              grossSales: '',
                              grossProfit: '',
                              netProfit: '',
                              totalAsset: '',
                              totalLiabilities: '',
                              netWorth: '',
                            },
                          ],
                        },
                      };
                    }, 'sec11Financials')
                  }
                  className="inline-flex items-center gap-1 px-2 py-1 text-[11px] font-medium text-emerald-900 bg-white border border-emerald-400 rounded cursor-pointer"
                >
                  <Plus className="w-3 h-3" /> Add Financial Row
                </button>
                <label className="block">
                  <span className="text-[11px] font-medium text-slate-700">Section 11 Summary / Note</span>
                  <textarea
                    rows={2}
                    value={currentSec11(aiSections.sec11Financials).summary}
                    onChange={(e) =>
                      onUpdateAiSections(
                        (prev) => ({
                          ...prev,
                          sec11Financials: { ...currentSec11(prev.sec11Financials), summary: e.target.value },
                        }),
                        'sec11Financials'
                      )
                    }
                    className="mt-0.5 w-full px-2 py-1 border border-slate-300 rounded bg-white"
                  />
                </label>
              </div>
            ) : (
              <>
                <table className="w-full table-fixed text-[10pt] border border-dotted border-black">
                  <colgroup>
                    <col className="w-[10%]" />
                    <col className="w-[15%]" />
                    <col className="w-[15%]" />
                    <col className="w-[15%]" />
                    <col className="w-[15%]" />
                    <col className="w-[15%]" />
                    <col className="w-[15%]" />
                  </colgroup>
                  <thead>
                    <tr>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Year</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Gross Sales</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Gross Profit</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Net Profit</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Total Asset</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Total liabilities</th>
                      <th className="border border-dotted border-black px-1 py-0.5 text-center">Net Worth</th>
                    </tr>
                  </thead>
                  <tbody>
                    {finRows.map((r, i) => (
                      <tr key={i} className={!r.year && !r.grossSales ? 'h-6' : ''}>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.year}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.grossSales}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.grossProfit}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.netProfit}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.totalAsset}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.totalLiabilities}</td>
                        <td className="border border-dotted border-black px-1 py-0.5 text-center">{r.netWorth}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </>
            )}
          </div>

          {/* SECTION 12: Contents of the Client's Prayer */}
          <div className="space-y-1 text-[10pt]">
            <div className="flex flex-wrap items-start justify-between gap-2">
              <p className="font-bold underline pt-0.5">12. Contents of the Client’s Prayer:</p>
              <SectionActionToolbar {...toolbarProps('clientPrayer')} />
            </div>
            {isEditing('clientPrayer') ? (
              <textarea
                rows={4}
                value={aiSections.clientPrayer ?? memo.clientPrayer}
                onChange={(e) => updateNarrativeAi('clientPrayer', e.target.value)}
                placeholder="Enter Contents of the Client’s Prayer or click Scrutinize and Fill Up Through AI..."
                className="w-full font-memo text-[10pt] p-2 border border-emerald-600 rounded bg-emerald-50/20 focus:outline-none"
              />
            ) : memo.clientPrayer ? (
              <p className="whitespace-pre-line">{memo.clientPrayer}</p>
            ) : null}
          </div>

          {/* SECTION 12: Branch Visit Report */}
          <div className="space-y-1 text-[10pt]">
            <div className="flex flex-wrap items-start justify-between gap-2">
              <p className="font-bold underline pt-0.5">12. Branch Visit Report:</p>
              <SectionActionToolbar {...toolbarProps('branchVisit')} />
            </div>
            {isEditing('branchVisit') ? (
              <textarea
                rows={5}
                value={aiSections.branchVisit ?? memo.branchVisitReport}
                onChange={(e) => updateNarrativeAi('branchVisit', e.target.value)}
                placeholder="Enter third-person Branch Visit Report or click Scrutinize and Fill Up Through AI..."
                className="w-full font-memo text-[10pt] p-2 border border-emerald-600 rounded bg-emerald-50/20 focus:outline-none"
              />
            ) : memo.branchVisitReport ? (
              <p className="whitespace-pre-line">{memo.branchVisitReport}</p>
            ) : null}
          </div>

          {/* SECTION 13: Recommendation of The Branch */}
          <div className="space-y-1 text-[10pt]">
            <div className="flex flex-wrap items-start justify-between gap-2">
              <p className="pt-0.5">
                <span className="font-bold underline">13. Recommendation of The Branch:</span> The Branch
                Recommended the following vide Proposal Ref: <Token ph={ph} code="#64#" /> dated{' '}
                <Token ph={ph} code="#65#" />
              </p>
              <SectionActionToolbar {...toolbarProps('sec13BranchRecommendation')} />
            </div>
            {isEditing('sec13BranchRecommendation') ? (
              <textarea
                rows={4}
                value={aiSections.sec13BranchRecommendation ?? memo.sec13BranchRecommendation}
                onChange={(e) => updateNarrativeAi('sec13BranchRecommendation', e.target.value)}
                placeholder="Enter Branch Recommendation summary or click Scrutinize and Fill Up Through AI..."
                className="w-full font-memo text-[10pt] p-2 border border-emerald-600 rounded bg-emerald-50/20 focus:outline-none"
              />
            ) : memo.sec13BranchRecommendation ? (
              <p className="pl-4 whitespace-pre-line">{memo.sec13BranchRecommendation}</p>
            ) : null}
          </div>

          {/* SECTION 14: Recommendation of Internal Control & Compliance Division */}
          <div className="space-y-1 text-[10pt]">
            <div className="flex flex-wrap items-start justify-between gap-2">
              <p className="pt-0.5">
                <span className="font-bold underline">
                  14. Recommendation of Internal Control &amp; Compliance Division:
                </span>{' '}
                {memo.sec14IccdRecommendation ? (
                  <span className="whitespace-pre-line">{memo.sec14IccdRecommendation}</span>
                ) : (
                  <span className="font-bold">N/A</span>
                )}
              </p>
              <SectionActionToolbar {...toolbarProps('sec14IccdRecommendation')} />
            </div>
            {isEditing('sec14IccdRecommendation') && (
              <textarea
                rows={3}
                value={aiSections.sec14IccdRecommendation ?? memo.sec14IccdRecommendation}
                onChange={(e) => updateNarrativeAi('sec14IccdRecommendation', e.target.value)}
                placeholder="Enter ICCD Recommendation or leave N/A..."
                className="w-full font-memo text-[10pt] p-2 border border-emerald-600 rounded bg-emerald-50/20 focus:outline-none"
              />
            )}
          </div>

          {/* SECTION 15: Observation and Recommendation of SAMD */}
          <div className="space-y-1 text-[10pt]">
            <div className="flex flex-wrap items-start justify-between gap-2">
              <p className="font-bold underline pt-0.5">15. Observation and Recommendation of SAMD:</p>
              <SectionActionToolbar {...toolbarProps('samdObservation')} />
            </div>
            {isEditing('samdObservation') ? (
              <textarea
                rows={7}
                value={aiSections.samdObservation ?? memo.samdObservation}
                onChange={(e) => updateNarrativeAi('samdObservation', e.target.value)}
                placeholder="Enter SAMD Observation and Recommendation points or click Scrutinize and Fill Up Through AI..."
                className="w-full font-memo text-[10pt] p-2 border border-emerald-600 rounded bg-emerald-50/20 focus:outline-none"
              />
            ) : memo.samdObservation ? (
              <div className="pl-4 whitespace-pre-line">{memo.samdObservation}</div>
            ) : null}
          </div>

          {/* Section 16 */}
          <div className="space-y-1">
            <p className="font-bold underline text-[10pt]">16. Particulars of the facility:</p>
            <table className="w-full table-fixed text-[10pt] border border-dotted border-black">
              <colgroup>
                <col className="w-[5%]" />
                <col className="w-[24%]" />
                <col className="w-[71%]" />
              </colgroup>
              <tbody>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5 text-center">a.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Facility</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    EXIT with Waiver under BRPD 1 Circular Letter No. 23 dated 29.06.2026
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">b.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Outstanding</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    Tk. <Token ph={ph} code="#25#" /> million as on <Token ph={ph} code="#68#" />.
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">c.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Down Payment</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    Down-payment of Tk. <Token ph={ph} code="#35#" /> Million is kept in{' '}
                    <Token ph={ph} code="#36#" /> A/C <Token ph={ph} code="#36#" />
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">d.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Waiver</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    Tk. <Token ph={ph} code="#4#" /> million as on <Token ph={ph} code="#68#" />
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">e.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Settlement/Exit amount</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    Tk. <Token ph={ph} code="#27#" /> Million as on <Token ph={ph} code="#68#" />
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">f.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Settlement period</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5 font-bold">
                    <Token ph={ph} code="#28#" />
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">g.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Profit/Rent rate</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">
                    Without charging further profit.
                  </td>
                </tr>
                <tr>
                  <td className="border border-dotted border-black px-1.5 py-0.5">h.</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Repayment Method</td>
                  <td className="border border-dotted border-black px-1.5 py-0.5">Lump sum.</td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* Section 17 */}
          <div className="space-y-1 text-[10pt]">
            <p className="font-bold underline text-[10pt]">17. Special Term(s) &amp; Condition(s):</p>
            <p className="pl-4">
              If the proposal approved, the branch is advised to comply the following special terms
              and conditions:
            </p>
            <ol className="pl-10 space-y-1.5 list-none leading-relaxed">
              <li>
                i. The waiver of profit shall be adjusted from FAD, Head Office from the provision
                amount kept (if any) and the rest amount shall be adjusted by branch gradually after
                adjustment of the settlement amount within the approved time frame.
              </li>
              <li>
                ii. The final settlement amount will be determined at the date of execution of
                settlement considering up to date liability position.
              </li>
              <li>
                iii. The Branch must ensure adjustment of all other direct and indirect liabilities
                of the client before allowing the final redemption of the mortgaged collateral
                security(s).
              </li>
              <li>
                iv. The Branch must ensure adjustment of all non-funded liabilities (L/C, Bank
                Guarantee, Credit Card etc.) before allowing final redemption of the mortgaged
                collateral security(s). The Branch must also ensure that any encashable security
                (Margin for Bank Guarantee, L/C, MTDR, Scheme deposit in the name of the client)
                including collateral is properly mentioned in the proposal. If any encashable
                security (Margin for Bank Guarantee, L/C, MTDR, Scheme deposit in the name of the
                client) is found not mentioned in the proposal and later encashed in favor of the
                client, the Manager will be held responsible.
              </li>
              <li>
                v. If the client fails to repay or fails to adjust the liability within the given
                time or fails to pay installments regularly, all facilities shall be cancelled and
                legal action shall be initiated by the Branch.
              </li>
              <li>
                vi. Branch should be utmost sincere to protect bank interest in the course of
                execution partial redemption.
              </li>
              <li>
                vii. Branch is advised to follow the guidelines of Bangladesh Bank BRPD circular No.
                06/22, Circular Letter No.18/22, 46/22 and subsequent up-dates as well as all
                instruction circular(s) of SIBL.
              </li>
              <li className="font-bold">
                viii. The Profit rate of the stipulated investment shall not be changed or reduced to
                0.00% until full and final adjustment of the Investment, as approved by Head Office.
              </li>
              <li>
                ix. All the legal expenses and other charges (if any) already incurred and to be
                incurred in connection with the withdrawal of case/suits and others issues shall be
                borne by the client.
              </li>
            </ol>
          </div>

          {/* Closing Box & Signatories */}
          <div className="pt-4 space-y-5 text-[10pt]">
            <div className="border border-black px-3 py-1 text-center font-bold">
              The proposal is hereby placed before the Investment Committee for kind recommendation.
            </div>

            <div className="font-bold leading-snug">
              <p>Union Bank PLC Unit</p>
              <p>Special Asset Management Division</p>
            </div>

            <p className="text-center font-bold">Submitted for kind decision.</p>

            <div className="grid grid-cols-2 gap-8 pt-6">
              <div>
                <p className="font-bold">(Shahidur Rahman khan)</p>
                <p>Deputy Managing Director &amp; CRO</p>
                <p>Sammilito Islami Bank PLC.</p>
              </div>
              <div>
                <p className="font-bold">(Ahmed Zafrul Hasan)</p>
                <p>Additional Managing Director &amp; COO</p>
                <p>Sammilito Islami Bank PLC.</p>
              </div>
            </div>

            <div className="pt-6 text-center leading-snug">
              <p className="font-bold">(Abedur Rahman Sikder)</p>
              <p>Managing Director</p>
              <p>Sammilito Islami Bank PLC.</p>
              <p>(Transferee of the assets and liabilities of</p>
              <p>EXIM Bank, SIBL, FSIB, Global Islami</p>
              <p>Bank and Union Bank)</p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
