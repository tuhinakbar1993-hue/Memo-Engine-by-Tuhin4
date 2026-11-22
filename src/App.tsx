import { useEffect, useMemo, useRef, useState } from 'react';
import { AlertTriangle, CheckCircle2, RefreshCw, Trash2 } from 'lucide-react';
import { ICMemoPreview } from './components/ICMemoPreview';
import { LandingScreen } from './components/LandingScreen';
import { ManualInputsTab } from './components/ManualInputsTab';
import { StageMemoTab } from './components/StageMemoTab';
import {
  INITIAL_STAGE_AI_SECTIONS,
  INITIAL_STAGE_INPUTS,
  resolveICMemo,
  resolveStageMemo,
  STAGE_SECTIONS_BY_DOC,
} from './engines/icMemoEngine';
import { inspectDocxTemplate, verifyPopulatedDocxBlob } from './exporters/docxTemplateEngine';
import { generatePdfFromDocxBlob, generateSynchronizedDocxBlob } from './exporters/pdfExporter';
import {
  detectDocumentKeyFromFile,
  parseAccountStatementPdf,
  parseBranchProposalFile,
  parseCibReportPdf,
  parseCostOfFundFile,
  parseOutstandingPdf,
  parseOverduePdf,
  parsePartyPerformancePdf,
  prepareBranchProposalSource,
  prepareMemoDocumentSource,
  warmBranchProposalSource,
} from './parsers/documentParsers';
import type {
  AiSectionsData,
  AppliedSyncState,
  BatchScrutinyState,
  BranchProposalData,
  BranchProposalSource,
  DocumentKey,
  DocumentUploadMeta,
  ExtractedDocumentsState,
  ManualInputsData,
  MemoDocType,
  MemoTabKey,
  MemoTemplateRecord,
  ScrutinySectionKey,
  StageAiSectionsData,
  StageBatchScrutinyState,
  StageMemoDocType,
  StageMemoInputs,
  StageScrutinySectionKey,
  StageSourceAuditResult,
  StageSourceUploadState,
} from './types/memo';
import {
  clearCentralCaseState,
  loadCentralCaseState,
  saveCentralCaseState,
} from './utils/caseStorage';
import {
  deleteTemplateRecord,
  loadAllTemplateRecords,
  saveTemplateRecord,
} from './utils/templateStorage';

function isHtmlPayload(contentType: string, text: string): boolean {
  if (/text\/html/i.test(contentType)) return true;
  const trimmed = text.trimStart();
  return /^<!doctype\b/i.test(trimmed) || /^<html\b/i.test(trimmed) || trimmed.startsWith('<');
}

function extractApiErrorFromRaw(
  rawText: string,
  contentType: string,
  status: number,
  fallbackLabel: string
): string {
  if (isHtmlPayload(contentType, rawText)) {
    if (status === 413) {
      return 'Uploaded document payload is too large for the server. Please upload a smaller file or fewer pages.';
    }
    if (status === 401 || status === 403) {
      return `Authentication or session check intercepted ${fallbackLabel} (${status}). Please reload the page and try again.`;
    }
    if (status === 502 || status === 503 || status === 504) {
      return `Server temporarily unavailable (${status}) while running ${fallbackLabel}. Please retry in a moment.`;
    }
    return `API returned an HTML page instead of JSON (${status}) for ${fallbackLabel}. Please reload the page or verify the backend server is running.`;
  }
  try {
    const parsed = JSON.parse(rawText) as { error?: unknown; message?: unknown };
    if (typeof parsed?.error === 'string' && parsed.error.trim()) {
      return parsed.error;
    }
    if (typeof parsed?.message === 'string' && parsed.message.trim()) {
      return parsed.message;
    }
  } catch {
    // Non-JSON plain text response
  }
  const plain = rawText.trim().split('\n')[0]?.slice(0, 200);
  return plain || `${fallbackLabel} (${status}).`;
}

async function parseJsonResponse<T>(response: Response, fallbackLabel: string): Promise<T> {
  const contentType = response.headers.get('content-type') ?? '';
  const rawText = await response.text();

  if (!response.ok) {
    throw new Error(extractApiErrorFromRaw(rawText, contentType, response.status, fallbackLabel));
  }

  if (!rawText.trim()) {
    throw new Error(`${fallbackLabel}: Server returned an empty response.`);
  }

  if (isHtmlPayload(contentType, rawText)) {
    throw new Error(
      `API endpoint returned an HTML page instead of JSON for ${fallbackLabel}. Please reload the page and retry.`
    );
  }

  try {
    return JSON.parse(rawText) as T;
  } catch {
    throw new Error(`${fallbackLabel}: Server returned an invalid JSON response.`);
  }
}

async function assertSseResponseOrThrow(response: Response, fallbackLabel: string): Promise<void> {
  const contentType = response.headers.get('content-type') ?? '';
  if (!response.ok) {
    const rawText = await response.text().catch(() => '');
    throw new Error(extractApiErrorFromRaw(rawText, contentType, response.status, fallbackLabel));
  }
  if (/text\/html/i.test(contentType)) {
    throw new Error(
      `API endpoint returned an HTML page instead of a streaming response for ${fallbackLabel}. Please reload the page and retry.`
    );
  }
}

const INITIAL_DOC_METAS: Record<DocumentKey, DocumentUploadMeta> = {
  branchProposal: {
    key: 'branchProposal',
    label: 'Branch Proposal',
    acceptedFormats: '.pdf,.docx,.xlsx,.xls,image/*',
    uploaded: false,
    fileName: '',
    parsing: false,
    stage: null,
    error: null,
  },
  outstanding: {
    key: 'outstanding',
    label: 'Outstanding',
    acceptedFormats: '.pdf,.docx,.xlsx,.xls',
    uploaded: false,
    fileName: '',
    parsing: false,
    stage: null,
    error: null,
  },
  partyPerformance: {
    key: 'partyPerformance',
    label: 'Party Performance',
    acceptedFormats: '.pdf,.docx,.xlsx,.xls',
    uploaded: false,
    fileName: '',
    parsing: false,
    stage: null,
    error: null,
  },
  overdue: {
    key: 'overdue',
    label: 'Overdue',
    acceptedFormats: '.pdf,.docx,.xlsx,.xls',
    uploaded: false,
    fileName: '',
    parsing: false,
    stage: null,
    error: null,
  },
  accountStatement: {
    key: 'accountStatement',
    label: 'Account Statement',
    acceptedFormats: '.pdf,.docx,.xlsx,.xls',
    uploaded: false,
    fileName: '',
    parsing: false,
    stage: null,
    error: null,
  },
  cibReport: {
    key: 'cibReport',
    label: 'CIB Report',
    acceptedFormats: '.pdf,.docx,.xlsx,.xls',
    uploaded: false,
    fileName: '',
    parsing: false,
    stage: null,
    error: null,
  },
  costOfFund: {
    key: 'costOfFund',
    label: 'Cost of Fund',
    acceptedFormats: '.xlsx,.xls,.csv,.docx,.pdf',
    uploaded: false,
    fileName: '',
    parsing: false,
    stage: null,
    error: null,
  },
};

const INITIAL_EXTRACTED_DOCS: ExtractedDocumentsState = {
  branchProposal: null,
  outstanding: null,
  partyPerformance: null,
  overdue: null,
  accountStatement: null,
  cibReport: null,
  costOfFund: null,
};

const INITIAL_MANUAL_INPUTS: ManualInputsData = {
  proposalDate: '',
  waiverAmount: '',
  settlementAmount: '',
  settlementPeriod: '',
  waiverRecClientBranch: '',
  waiverRecHO: '',
  sourceOfRepayment: '',
};

const INITIAL_AI_SECTIONS: AiSectionsData = {
  justification: null,
  sec1NatureOfBusiness: null,
  sec2LiabilityOurBank: null,
  sec3LiabilityOtherUnits: null,
  sec4CibStatus: null,
  sec5LegalAction: null,
  sec6AuditCompliance: null,
  sec7Security: null,
  sec8SourceOfRepayment: null,
  sec9HoSanctions: null,
  sec10ClientParticulars: null,
  sec11Financials: null,
  clientPrayer: null,
  branchVisit: null,
  sec13BranchRecommendation: null,
  sec14IccdRecommendation: null,
  samdObservation: null,
};

const TABS: Array<{ id: MemoTabKey; label: string }> = [
  { id: 'manual', label: 'Manual Inputs' },
  { id: 'ic', label: 'IC Memo' },
  { id: 'mec', label: 'MEC Memo' },
  { id: 'board', label: 'Board Memo' },
  { id: 'sanction', label: 'Sanction' },
];

const MANUAL_FIELD_TO_CODE: Partial<Record<keyof ManualInputsData, string>> = {
  proposalDate: '#1#',
  waiverAmount: '#4#',
  settlementAmount: '#27#',
  settlementPeriod: '#28#',
  waiverRecClientBranch: '#45#',
  waiverRecHO: '#47#',
};

const ALL_SCRUTINY_SECTIONS: ScrutinySectionKey[] = [
  'justification',
  'sec1NatureOfBusiness',
  'sec2LiabilityOurBank',
  'sec3LiabilityOtherUnits',
  'sec4CibStatus',
  'sec5LegalAction',
  'sec6AuditCompliance',
  'sec7Security',
  'sec8SourceOfRepayment',
  'sec9HoSanctions',
  'sec10ClientParticulars',
  'sec11Financials',
  'clientPrayer',
  'branchVisit',
  'sec13BranchRecommendation',
  'sec14IccdRecommendation',
  'samdObservation',
];

const INITIAL_BATCH_SCRUTINY: BatchScrutinyState = {
  running: false,
  total: 0,
  targetSections: [],
  completedSections: [],
  retryingSections: [],
  failedSections: [],
  insufficientSections: [],
  skippedManualSections: [],
};

const INITIAL_STAGE_BATCH_SCRUTINY: StageBatchScrutinyState = {
  running: false,
  docType: null,
  total: 0,
  targetSections: [],
  completedSections: [],
  retryingSections: [],
  failedSections: [],
  insufficientSections: [],
  skippedManualSections: [],
};

const DOC_FILE_PREFIX: Record<MemoDocType, string> = {
  ic: 'IC_Memo',
  mec: 'MEC_Memo',
  board: 'Board_Memo',
  sanction: 'Sanction',
};

const createInitialStageSourceUploads = (): Record<StageMemoDocType, StageSourceUploadState> => ({
  mec: {
    uploaded: false,
    fileName: '',
    sourceType: 'ic_memo',
    processing: false,
    stage: null,
    error: null,
    extractedAt: null,
    extractedFieldCount: 0,
    audit: null,
  },
  board: {
    uploaded: false,
    fileName: '',
    sourceType: 'mec_memo',
    processing: false,
    stage: null,
    error: null,
    extractedAt: null,
    extractedFieldCount: 0,
    audit: null,
  },
  sanction: {
    uploaded: false,
    fileName: '',
    sourceType: 'board_memo',
    processing: false,
    stage: null,
    error: null,
    extractedAt: null,
    extractedFieldCount: 0,
    audit: null,
  },
});

export default function App() {
  const [entered, setEntered] = useState(false);
  const [activeTab, setActiveTab] = useState<MemoTabKey>('manual');

  const [docMetas, setDocMetas] = useState<Record<DocumentKey, DocumentUploadMeta>>(INITIAL_DOC_METAS);
  const [extractedDocs, setExtractedDocs] = useState<ExtractedDocumentsState>(INITIAL_EXTRACTED_DOCS);
  const [manualInputs, setManualInputs] = useState<ManualInputsData>(INITIAL_MANUAL_INPUTS);
  const [aiSections, setAiSections] = useState<AiSectionsData>(INITIAL_AI_SECTIONS);
  const [overrides, setOverrides] = useState<Record<string, string>>({});

  const [templates, setTemplates] = useState<Record<MemoDocType, MemoTemplateRecord | null>>({
    ic: null,
    mec: null,
    board: null,
    sanction: null,
  });

  const [loadingSamples, setLoadingSamples] = useState(false);
  const [batchProgress, setBatchProgress] = useState<{
    active: boolean;
    total: number;
    completed: number;
  }>({ active: false, total: 0, completed: 0 });
  const [branchProposalFile, setBranchProposalFile] = useState<File | null>(null);
  const [scrutinizingSection, setScrutinizingSection] = useState<ScrutinySectionKey | null>(null);
  const [batchScrutiny, setBatchScrutiny] = useState<BatchScrutinyState>(INITIAL_BATCH_SCRUTINY);
  const [manualEditedSections, setManualEditedSections] = useState<
    Partial<Record<ScrutinySectionKey, boolean>>
  >({});
  const [stageInputs, setStageInputs] = useState<StageMemoInputs>(INITIAL_STAGE_INPUTS);
  const [stageAiSections, setStageAiSections] = useState<StageAiSectionsData>(
    INITIAL_STAGE_AI_SECTIONS
  );
  const [stageManualEditedSections, setStageManualEditedSections] = useState<
    Partial<Record<StageScrutinySectionKey, boolean>>
  >({});
  const [scrutinizingStageSection, setScrutinizingStageSection] =
    useState<StageScrutinySectionKey | null>(null);
  const [stageBatchScrutiny, setStageBatchScrutiny] = useState<StageBatchScrutinyState>(
    INITIAL_STAGE_BATCH_SCRUTINY
  );
  const [stageSourceUploads, setStageSourceUploads] = useState<
    Record<StageMemoDocType, StageSourceUploadState>
  >(createInitialStageSourceUploads);
  const [showClearConfirm, setShowClearConfirm] = useState(false);
  const [clearingCaseData, setClearingCaseData] = useState(false);
  const [aiError, setAiError] = useState<string | null>(null);
  const [downloading, setDownloading] = useState(false);
  const [downloadingPdf, setDownloadingPdf] = useState(false);

  const serverCachedSourceIdRef = useRef<string | null>(null);
  const stageSourceFilesRef = useRef<Partial<Record<StageMemoDocType, File>>>({});
  const stageServerCachedSourceIdsRef = useRef<Partial<Record<StageMemoDocType, string>>>({});
  const cacheWarmPromiseRef = useRef<{ sourceId?: string; promise: Promise<void> } | null>(null);
  const batchAbortRef = useRef<AbortController | null>(null);
  const stageBatchAbortRef = useRef<AbortController | null>(null);
  const caseDirtyRef = useRef(false);
  const persistTimerRef = useRef<number | null>(null);

  const caseRef = useRef<{
    extractedDocs: ExtractedDocumentsState;
    manualInputs: ManualInputsData;
    aiSections: AiSectionsData;
    overrides: Record<string, string>;
    manualEditedSections: Partial<Record<ScrutinySectionKey, boolean>>;
    stageInputs: StageMemoInputs;
    stageAiSections: StageAiSectionsData;
    stageManualEditedSections: Partial<Record<StageScrutinySectionKey, boolean>>;
    stageSourceUploads: Record<StageMemoDocType, StageSourceUploadState>;
  }>({
    extractedDocs: INITIAL_EXTRACTED_DOCS,
    manualInputs: INITIAL_MANUAL_INPUTS,
    aiSections: INITIAL_AI_SECTIONS,
    overrides: {},
    manualEditedSections: {},
    stageInputs: INITIAL_STAGE_INPUTS,
    stageAiSections: INITIAL_STAGE_AI_SECTIONS,
    stageManualEditedSections: {},
    stageSourceUploads: createInitialStageSourceUploads(),
  });
  const docMetasRef = useRef<Record<DocumentKey, DocumentUploadMeta>>(INITIAL_DOC_METAS);

  const resolvedICMemo = useMemo(
    () =>
      resolveICMemo(
        extractedDocs,
        manualInputs,
        aiSections,
        overrides,
        manualEditedSections
      ),
    [extractedDocs, manualInputs, aiSections, overrides, manualEditedSections]
  );

  const resolvedMecMemo = useMemo(
    () =>
      resolveStageMemo(
        'mec',
        resolvedICMemo,
        stageInputs,
        stageAiSections,
        stageManualEditedSections
      ),
    [resolvedICMemo, stageInputs, stageAiSections, stageManualEditedSections]
  );

  const resolvedBoardMemo = useMemo(
    () =>
      resolveStageMemo(
        'board',
        resolvedICMemo,
        stageInputs,
        stageAiSections,
        stageManualEditedSections
      ),
    [resolvedICMemo, stageInputs, stageAiSections, stageManualEditedSections]
  );

  const resolvedSanctionMemo = useMemo(
    () =>
      resolveStageMemo(
        'sanction',
        resolvedICMemo,
        stageInputs,
        stageAiSections,
        stageManualEditedSections
      ),
    [resolvedICMemo, stageInputs, stageAiSections, stageManualEditedSections]
  );

  const [appliedSnapshot, setAppliedSnapshot] = useState<AppliedSyncState>(() => ({
    appliedAt: null,
    version: 0,
    memo: resolveICMemo(INITIAL_EXTRACTED_DOCS, INITIAL_MANUAL_INPUTS, INITIAL_AI_SECTIONS, {}, {}),
    stageInputs: INITIAL_STAGE_INPUTS,
    stageAiSections: INITIAL_STAGE_AI_SECTIONS,
  }));
  const [hasPendingChanges, setHasPendingChanges] = useState(false);
  const [applySignal, setApplySignal] = useState(0);
  const [applyNotice, setApplyNotice] = useState<string | null>(null);

  useEffect(() => {
    caseRef.current = {
      extractedDocs,
      manualInputs,
      aiSections,
      overrides,
      manualEditedSections,
      stageInputs,
      stageAiSections,
      stageManualEditedSections,
      stageSourceUploads,
    };
  }, [
    extractedDocs,
    manualInputs,
    aiSections,
    overrides,
    manualEditedSections,
    stageInputs,
    stageAiSections,
    stageManualEditedSections,
    stageSourceUploads,
  ]);

  useEffect(() => {
    docMetasRef.current = docMetas;
  }, [docMetas]);

  const serializeCaseForStorage = () => {
    const cur = caseRef.current;
    const cleanUploads: Partial<Record<StageMemoDocType, StageSourceUploadState>> = {};
    for (const k of ['mec', 'board', 'sanction'] as StageMemoDocType[]) {
      cleanUploads[k] = {
        ...cur.stageSourceUploads[k],
        processing: false,
        stage: null,
      };
    }
    return {
      ...cur,
      stageSourceUploads: cleanUploads,
    };
  };

  const persistCentralCaseNow = async (immediate = false) => {
    if (persistTimerRef.current !== null) {
      window.clearTimeout(persistTimerRef.current);
      persistTimerRef.current = null;
    }
    if (immediate) {
      await saveCentralCaseState(serializeCaseForStorage(), docMetasRef.current);
      return;
    }
    persistTimerRef.current = window.setTimeout(() => {
      persistTimerRef.current = null;
      void saveCentralCaseState(serializeCaseForStorage(), docMetasRef.current);
    }, 150);
  };

  const syncAppliedSnapshotFromRef = () => {
    const latest = caseRef.current;
    const syncedMemo = resolveICMemo(
      latest.extractedDocs,
      latest.manualInputs,
      latest.aiSections,
      latest.overrides,
      latest.manualEditedSections
    );
    const nowStr = new Date().toLocaleTimeString([], {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
    });
    setAppliedSnapshot((prev) => ({
      appliedAt: nowStr,
      version: prev.version + 1,
      memo: syncedMemo,
      stageInputs: latest.stageInputs,
      stageAiSections: latest.stageAiSections,
    }));
    setHasPendingChanges(false);
    return syncedMemo;
  };

  const setDocMetasSync = (
    updater:
      | Record<DocumentKey, DocumentUploadMeta>
      | ((prev: Record<DocumentKey, DocumentUploadMeta>) => Record<DocumentKey, DocumentUploadMeta>)
  ) => {
    const next = typeof updater === 'function' ? updater(docMetasRef.current) : updater;
    docMetasRef.current = next;
    setDocMetas(next);
  };

  const setExtractedDocsSync = (
    updater:
      | ExtractedDocumentsState
      | ((prev: ExtractedDocumentsState) => ExtractedDocumentsState)
  ) => {
    caseDirtyRef.current = true;
    const next = typeof updater === 'function' ? updater(caseRef.current.extractedDocs) : updater;
    caseRef.current = { ...caseRef.current, extractedDocs: next };
    setExtractedDocs(next);
  };

  const setManualInputsSync = (
    updater: ManualInputsData | ((prev: ManualInputsData) => ManualInputsData)
  ) => {
    caseDirtyRef.current = true;
    const next = typeof updater === 'function' ? updater(caseRef.current.manualInputs) : updater;
    caseRef.current = { ...caseRef.current, manualInputs: next };
    setManualInputs(next);
  };

  const setAiSectionsSync = (
    updater: AiSectionsData | ((prev: AiSectionsData) => AiSectionsData)
  ) => {
    caseDirtyRef.current = true;
    const next = typeof updater === 'function' ? updater(caseRef.current.aiSections) : updater;
    caseRef.current = { ...caseRef.current, aiSections: next };
    setAiSections(next);
  };

  const setOverridesSync = (
    updater: Record<string, string> | ((prev: Record<string, string>) => Record<string, string>)
  ) => {
    caseDirtyRef.current = true;
    const next = typeof updater === 'function' ? updater(caseRef.current.overrides) : updater;
    caseRef.current = { ...caseRef.current, overrides: next };
    setOverrides(next);
  };

  const setManualEditedSectionsSync = (
    updater:
      | Partial<Record<ScrutinySectionKey, boolean>>
      | ((
          prev: Partial<Record<ScrutinySectionKey, boolean>>
        ) => Partial<Record<ScrutinySectionKey, boolean>>)
  ) => {
    caseDirtyRef.current = true;
    const next =
      typeof updater === 'function' ? updater(caseRef.current.manualEditedSections) : updater;
    caseRef.current = { ...caseRef.current, manualEditedSections: next };
    setManualEditedSections(next);
  };

  const setStageInputsSync = (
    updater: StageMemoInputs | ((prev: StageMemoInputs) => StageMemoInputs)
  ) => {
    caseDirtyRef.current = true;
    const next = typeof updater === 'function' ? updater(caseRef.current.stageInputs) : updater;
    caseRef.current = { ...caseRef.current, stageInputs: next };
    setStageInputs(next);
  };

  const setStageAiSectionsSync = (
    updater: StageAiSectionsData | ((prev: StageAiSectionsData) => StageAiSectionsData)
  ) => {
    caseDirtyRef.current = true;
    const next = typeof updater === 'function' ? updater(caseRef.current.stageAiSections) : updater;
    caseRef.current = { ...caseRef.current, stageAiSections: next };
    setStageAiSections(next);
  };

  const setStageManualEditedSectionsSync = (
    updater:
      | Partial<Record<StageScrutinySectionKey, boolean>>
      | ((
          prev: Partial<Record<StageScrutinySectionKey, boolean>>
        ) => Partial<Record<StageScrutinySectionKey, boolean>>)
  ) => {
    caseDirtyRef.current = true;
    const next =
      typeof updater === 'function' ? updater(caseRef.current.stageManualEditedSections) : updater;
    caseRef.current = { ...caseRef.current, stageManualEditedSections: next };
    setStageManualEditedSections(next);
  };

  useEffect(() => {
    void loadAllTemplateRecords().then((loaded) => {
      setTemplates(loaded);
    });
    void loadCentralCaseState().then((saved) => {
      if (!saved || caseDirtyRef.current) return;
      const defaultStageUploads = createInitialStageSourceUploads();
      const hydratedStageUploads: Record<StageMemoDocType, StageSourceUploadState> = {
        mec: {
          ...defaultStageUploads.mec,
          ...(saved.stageSourceUploads?.mec || {}),
          processing: false,
          stage: null,
        },
        board: {
          ...defaultStageUploads.board,
          ...(saved.stageSourceUploads?.board || {}),
          processing: false,
          stage: null,
        },
        sanction: {
          ...defaultStageUploads.sanction,
          ...(saved.stageSourceUploads?.sanction || {}),
          processing: false,
          stage: null,
        },
      };

      const hydrated = {
        extractedDocs: { ...INITIAL_EXTRACTED_DOCS, ...(saved.extractedDocs || {}) },
        manualInputs: { ...INITIAL_MANUAL_INPUTS, ...(saved.manualInputs || {}) },
        aiSections: { ...INITIAL_AI_SECTIONS, ...(saved.aiSections || {}) },
        overrides: saved.overrides || {},
        manualEditedSections: saved.manualEditedSections || {},
        stageInputs: { ...INITIAL_STAGE_INPUTS, ...(saved.stageInputs || {}) },
        stageAiSections: {
          mecIcRecommendation: saved.stageAiSections?.mecIcRecommendation ?? null,
          boardMecRecommendation: saved.stageAiSections?.boardMecRecommendation ?? null,
          sanctionCrossVerifiedSummary: saved.stageAiSections?.sanctionCrossVerifiedSummary ?? null,
        },
        stageManualEditedSections: {
          ...(saved.stageManualEditedSections?.mecIcRecommendation !== undefined
            ? { mecIcRecommendation: saved.stageManualEditedSections.mecIcRecommendation }
            : {}),
          ...(saved.stageManualEditedSections?.boardMecRecommendation !== undefined
            ? { boardMecRecommendation: saved.stageManualEditedSections.boardMecRecommendation }
            : {}),
          ...(saved.stageManualEditedSections?.sanctionCrossVerifiedSummary !== undefined
            ? { sanctionCrossVerifiedSummary: saved.stageManualEditedSections.sanctionCrossVerifiedSummary }
            : {}),
        },
        stageSourceUploads: hydratedStageUploads,
      };
      caseRef.current = hydrated;
      setExtractedDocs(hydrated.extractedDocs);
      setManualInputs(hydrated.manualInputs);
      setAiSections(hydrated.aiSections);
      setOverrides(hydrated.overrides);
      setManualEditedSections(hydrated.manualEditedSections);
      setStageInputs(hydrated.stageInputs);
      setStageAiSections(hydrated.stageAiSections);
      setStageManualEditedSections(hydrated.stageManualEditedSections);
      setStageSourceUploads(hydrated.stageSourceUploads);

      if (saved.docUploadedFlags) {
        setDocMetasSync((prev) => {
          const next = { ...prev };
          for (const key of Object.keys(INITIAL_DOC_METAS) as DocumentKey[]) {
            const flag = saved.docUploadedFlags?.[key];
            if (flag) {
              next[key] = {
                ...next[key],
                uploaded: Boolean(flag.uploaded),
                fileName: flag.fileName || '',
              };
            }
          }
          return next;
        });
      }

      const hydratedMemo = resolveICMemo(
        hydrated.extractedDocs,
        hydrated.manualInputs,
        hydrated.aiSections,
        hydrated.overrides,
        hydrated.manualEditedSections
      );
      setAppliedSnapshot({
        appliedAt: saved.updatedAt
          ? new Date(saved.updatedAt).toLocaleTimeString([], {
              hour: '2-digit',
              minute: '2-digit',
              second: '2-digit',
            })
          : null,
        version: 1,
        memo: hydratedMemo,
        stageInputs: hydrated.stageInputs,
        stageAiSections: hydrated.stageAiSections,
      });
      setHasPendingChanges(false);
    });

    return () => {
      if (persistTimerRef.current !== null) {
        window.clearTimeout(persistTimerRef.current);
      }
    };
  }, []);

  useEffect(() => {
    if (!applyNotice) return;
    const timer = window.setTimeout(() => setApplyNotice(null), 5000);
    return () => window.clearTimeout(timer);
  }, [applyNotice]);

  const updateDocStage = (key: DocumentKey, stage: string) => {
    setDocMetasSync((prev) => ({
      ...prev,
      [key]: {
        ...prev[key],
        stage,
      },
    }));
  };

  const handleFileUpload = async (key: DocumentKey, file: File) => {
    setDocMetasSync((prev) => ({
      ...prev,
      [key]: {
        ...prev[key],
        parsing: true,
        stage: 'Starting...',
        error: null,
        fileName: file.name,
      },
    }));

    const onStage = (stage: string) => updateDocStage(key, stage);

    try {
      if (key === 'branchProposal') {
        const bpData = await parseBranchProposalFile(file, onStage);
        setBranchProposalFile(file);
        setExtractedDocsSync((prev) => ({ ...prev, branchProposal: bpData }));
        serverCachedSourceIdRef.current = null;
        setBatchScrutiny(INITIAL_BATCH_SCRUTINY);
        // Pre-warm client page images + register server cache silently so AI calls start immediately
        const warmPromise = prepareBranchProposalSource(file)
          .then(async (src) => {
            if (!src.sourceId) return;
            if (cacheWarmPromiseRef.current?.promise === warmPromise) {
              cacheWarmPromiseRef.current.sourceId = src.sourceId;
            }
            const resp = await fetch('/api/gemini/cache-source', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ sourceId: src.sourceId, branchProposalSource: src }),
            });
            if (
              resp.ok &&
              (resp.headers.get('content-type') ?? '').toLowerCase().includes('application/json')
            ) {
              serverCachedSourceIdRef.current = src.sourceId;
            }
          })
          .catch(() => {
            // Silent pre-warm fallback; on-demand scrutiny sends source directly if needed
          });
        cacheWarmPromiseRef.current = { promise: warmPromise };
      } else if (key === 'outstanding') {
        const res = await parseOutstandingPdf(file, onStage);
        setExtractedDocsSync((prev) => ({ ...prev, outstanding: res.data }));
      } else if (key === 'partyPerformance') {
        const res = await parsePartyPerformancePdf(file, onStage);
        setExtractedDocsSync((prev) => ({ ...prev, partyPerformance: res.data }));
      } else if (key === 'overdue') {
        const res = await parseOverduePdf(file, onStage);
        setExtractedDocsSync((prev) => ({ ...prev, overdue: res.data }));
      } else if (key === 'accountStatement') {
        const res = await parseAccountStatementPdf(file, onStage);
        setExtractedDocsSync((prev) => ({ ...prev, accountStatement: res.data }));
      } else if (key === 'cibReport') {
        const res = await parseCibReportPdf(file, onStage);
        setExtractedDocsSync((prev) => ({ ...prev, cibReport: res.data }));
      } else if (key === 'costOfFund') {
        const res = await parseCostOfFundFile(file, onStage);
        setExtractedDocsSync((prev) => ({ ...prev, costOfFund: res.data }));
      }

      setDocMetasSync((prev) => ({
        ...prev,
        [key]: {
          ...prev[key],
          uploaded: true,
          parsing: false,
          stage: null,
          error: null,
          fileName: file.name,
        },
      }));
      syncAppliedSnapshotFromRef();
      void persistCentralCaseNow(true);
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Failed to parse document.';
      setDocMetasSync((prev) => ({
        ...prev,
        [key]: {
          ...prev[key],
          uploaded: false,
          parsing: false,
          stage: null,
          error: msg,
        },
      }));
    }
  };

  const handleLoadSampleDocuments = async () => {
    setLoadingSamples(true);
    const sampleFiles: Array<{ key: DocumentKey; url: string; name: string }> = [
      { key: 'outstanding', url: '/assets/aistudio/Outstanding_18.11.26_644e801e.pdf', name: 'Outstanding_18.11.26.pdf' },
      { key: 'partyPerformance', url: '/assets/aistudio/Party_Performance_55885ce3.pdf', name: 'Party_Performance.pdf' },
      { key: 'overdue', url: '/assets/aistudio/Overdue_a9f320c0.pdf', name: 'Overdue.pdf' },
      { key: 'accountStatement', url: '/assets/aistudio/Account_Statement_4197efed.pdf', name: 'Account_Statement.pdf' },
      { key: 'cibReport', url: '/assets/aistudio/CIB_Report_dec9ba1f.pdf', name: 'CIB_Report.pdf' },
      { key: 'costOfFund', url: '/assets/aistudio/Cost_of_Fund_3c8b0d9a.pdf', name: 'Cost_of_Fund.pdf' },
      { key: 'branchProposal', url: '/assets/aistudio/Branch_Proposal_a3692300.pdf', name: 'Branch_Proposal.pdf' },
    ];

    setBatchProgress({ active: true, total: sampleFiles.length, completed: 0 });

    try {
      await Promise.all(
        sampleFiles.map(async ({ key, url, name }) => {
          updateDocStage(key, 'Fetching sample...');
          try {
            const response = await fetch(url);
            const sampleContentType = response.headers.get('content-type') ?? '';
            if (!response.ok || /text\/html/i.test(sampleContentType)) {
              throw new Error(`Sample file missing: ${name}`);
            }
            const blob = await response.blob();
            const file = new File([blob], name, { type: blob.type || 'application/pdf' });
            await handleFileUpload(key, file);
          } finally {
            setBatchProgress((prev) => ({ ...prev, completed: prev.completed + 1 }));
          }
        })
      );
    } finally {
      setLoadingSamples(false);
      setBatchProgress({ active: false, total: 0, completed: 0 });
    }
  };

  const handleBatchUpload = async (files: FileList | File[]) => {
    const fileList = Array.from(files);
    const unmatched: string[] = [];

    setBatchProgress({ active: true, total: fileList.length, completed: 0 });

    try {
      await Promise.all(
        fileList.map(async (file) => {
          try {
            const key = await detectDocumentKeyFromFile(file);
            if (key) {
              await handleFileUpload(key, file);
            } else {
              unmatched.push(file.name);
            }
          } finally {
            setBatchProgress((prev) => ({ ...prev, completed: prev.completed + 1 }));
          }
        })
      );
    } finally {
      setBatchProgress({ active: false, total: 0, completed: 0 });
    }

    return unmatched;
  };

  const handleManualChange = (field: keyof ManualInputsData, value: string) => {
    setManualInputsSync((prev) => ({ ...prev, [field]: value }));
    if (field === 'sourceOfRepayment') {
      setManualEditedSectionsSync((prev) => ({
        ...prev,
        sec8SourceOfRepayment: value.trim() !== '',
      }));
    }
    const mappedCode = MANUAL_FIELD_TO_CODE[field];
    if (mappedCode) {
      setOverridesSync((prev) => {
        if (prev[mappedCode] === undefined) return prev;
        const next = { ...prev };
        delete next[mappedCode];
        return next;
      });
    }
    syncAppliedSnapshotFromRef();
    void persistCentralCaseNow(false);
  };

  const handleManualSourceOfRepaymentChange = (value: string) => {
    setManualInputsSync((prev) => ({ ...prev, sourceOfRepayment: value }));
    setAiSectionsSync((prev) => ({ ...prev, sec8SourceOfRepayment: value }));
    setManualEditedSectionsSync((prev) => ({ ...prev, sec8SourceOfRepayment: true }));
    syncAppliedSnapshotFromRef();
    void persistCentralCaseNow(false);
  };

  const handleOverrideChange = (code: string, value: string) => {
    setOverridesSync((prev) => ({ ...prev, [code]: value }));
    if (value.trim() !== '') {
      if (code === '#51#') {
        setManualEditedSectionsSync((prev) => ({ ...prev, justification: true }));
      } else if (code === '#54#') {
        setManualEditedSectionsSync((prev) => ({ ...prev, sec1NatureOfBusiness: true }));
      } else if (code === '#56#' || code === '#58#') {
        setManualEditedSectionsSync((prev) => ({ ...prev, sec2LiabilityOurBank: true }));
      } else if (code === '#62#') {
        setManualEditedSectionsSync((prev) => ({ ...prev, sec4CibStatus: true }));
      }
    }
    syncAppliedSnapshotFromRef();
    void persistCentralCaseNow(false);
  };

  const handleResetOverrides = () => {
    setOverridesSync({});
    syncAppliedSnapshotFromRef();
    void persistCentralCaseNow(true);
  };

  const clearOverridesForSection = (section: ScrutinySectionKey) => {
    const codesToClear =
      section === 'justification'
        ? ['#51#']
        : section === 'sec1NatureOfBusiness'
          ? ['#54#']
          : section === 'sec2LiabilityOurBank'
            ? ['#56#', '#58#']
            : section === 'sec4CibStatus'
              ? ['#62#']
              : [];
    if (codesToClear.length === 0) return;
    setOverridesSync((prev) => {
      let changed = false;
      const next = { ...prev };
      for (const c of codesToClear) {
        if (next[c] !== undefined) {
          delete next[c];
          changed = true;
        }
      }
      return changed ? next : prev;
    });
  };

  const buildScrutinyContext = (mode: 'ic' | 'stage' = 'stage') => {
    const latest = caseRef.current;
    const latestMemo = resolveICMemo(
      latest.extractedDocs,
      latest.manualInputs,
      latest.aiSections,
      latest.overrides,
      latest.manualEditedSections
    );
    const extractedSummaries = {
      outstanding: latest.extractedDocs.outstanding,
      partyPerformance: latest.extractedDocs.partyPerformance,
      overdue: latest.extractedDocs.overdue,
      accountStatement: latest.extractedDocs.accountStatement,
      cibReport: latest.extractedDocs.cibReport,
      costOfFund: latest.extractedDocs.costOfFund,
    };

    if (mode === 'ic') {
      return {
        placeholders: latestMemo.placeholders,
        sourceOfRepayment: latestMemo.sourceOfRepayment,
        clientPrayer: latestMemo.clientPrayer,
        branchVisitReport: latestMemo.branchVisitReport,
        justificationLines: latestMemo.justificationLines,
        sec13BranchRecommendation: latestMemo.sec13BranchRecommendation,
        sec14IccdRecommendation: latestMemo.sec14IccdRecommendation,
        extractedSummaries,
      };
    }

    const { branchProposal: _omitNestedBp, ...icMemoFinalizedClean } = latestMemo;
    return {
      placeholders: latestMemo.placeholders,
      sourceOfRepayment: latestMemo.sourceOfRepayment,
      clientPrayer: latestMemo.clientPrayer,
      branchVisitReport: latestMemo.branchVisitReport,
      justificationLines: latestMemo.justificationLines,
      sec13BranchRecommendation: latestMemo.sec13BranchRecommendation,
      sec14IccdRecommendation: latestMemo.sec14IccdRecommendation,
      samdObservation: latestMemo.samdObservation,
      sec2NonFundedNote: latestMemo.sec2NonFundedNote,
      sec7OthersNote: latestMemo.sec7OthersNote,
      sectionSummaries: latestMemo.sectionSummaries,
      icMemoFinalized: icMemoFinalizedClean,
      icMemoAiSections: latest.aiSections,
      extractedSummaries,
      stageInputs: latest.stageInputs,
      stageSections: latest.stageAiSections,
    };
  };

  const isNarrativeScrutinySection = (section: ScrutinySectionKey): boolean =>
    section === 'justification' ||
    section === 'sec1NatureOfBusiness' ||
    section === 'sec8SourceOfRepayment' ||
    section === 'clientPrayer' ||
    section === 'branchVisit' ||
    section === 'sec13BranchRecommendation' ||
    section === 'sec14IccdRecommendation' ||
    section === 'samdObservation';

  const applyScrutinizedSectionsBatch = (
    items: Array<{ section: ScrutinySectionKey; result: Record<string, unknown> }>,
    immediatePersist = true
  ) => {
    if (items.length === 0) return;

    let clearSec8Manual = false;
    const codesToClear: string[] = [];

    setAiSectionsSync((prev) => {
      const next = { ...prev };
      for (const { section, result } of items) {
        if (isNarrativeScrutinySection(section)) {
          (next as Record<string, unknown>)[section] = String(result.text ?? '').trim();
          if (section === 'sec8SourceOfRepayment') {
            clearSec8Manual = true;
          }
        } else {
          (next as Record<string, unknown>)[section] = result;
        }
        if (section === 'justification') codesToClear.push('#51#');
        else if (section === 'sec1NatureOfBusiness') codesToClear.push('#54#');
        else if (section === 'sec2LiabilityOurBank') codesToClear.push('#56#', '#58#');
        else if (section === 'sec4CibStatus') codesToClear.push('#62#');
      }
      return next;
    });

    if (clearSec8Manual) {
      setManualInputsSync((prev) =>
        prev.sourceOfRepayment === '' ? prev : { ...prev, sourceOfRepayment: '' }
      );
    }

    if (codesToClear.length > 0) {
      setOverridesSync((prev) => {
        let changed = false;
        const next = { ...prev };
        for (const c of codesToClear) {
          if (next[c] !== undefined) {
            delete next[c];
            changed = true;
          }
        }
        return changed ? next : prev;
      });
    }

    setManualEditedSectionsSync((prev) => {
      const next = { ...prev };
      for (const { section } of items) {
        next[section] = false;
      }
      return next;
    });

    syncAppliedSnapshotFromRef();
    void persistCentralCaseNow(immediatePersist);
  };

  const applyScrutinizedSectionResult = (
    section: ScrutinySectionKey,
    result: Record<string, unknown>
  ) => {
    applyScrutinizedSectionsBatch([{ section, result }], true);
  };

  const fetchWithSourceCache = async (
    url: string,
    payload: Record<string, unknown>,
    source: BranchProposalSource,
    signal?: AbortSignal
  ): Promise<Response> => {
    if (
      source.sourceId &&
      serverCachedSourceIdRef.current !== source.sourceId &&
      cacheWarmPromiseRef.current &&
      (!cacheWarmPromiseRef.current.sourceId ||
        cacheWarmPromiseRef.current.sourceId === source.sourceId)
    ) {
      await cacheWarmPromiseRef.current.promise.catch(() => undefined);
    }

    const canReuseServerCache =
      Boolean(source.sourceId) && serverCachedSourceIdRef.current === source.sourceId;

    const firstResponse = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      signal,
      body: JSON.stringify(
        canReuseServerCache
          ? { ...payload, sourceId: source.sourceId }
          : { ...payload, sourceId: source.sourceId, branchProposalSource: source }
      ),
    });

    if (firstResponse.status === 404 && canReuseServerCache) {
      serverCachedSourceIdRef.current = null;
      return fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        signal,
        body: JSON.stringify({
          ...payload,
          sourceId: source.sourceId,
          branchProposalSource: source,
        }),
      });
    }

    return firstResponse;
  };

  const handleScrutinizeSection = async (
    section: ScrutinySectionKey,
    forceOverwrite = false
  ) => {
    if (batchScrutiny.running) return;
    if (!branchProposalFile) {
      setAiError(
        'Branch Proposal is not uploaded yet. Please upload the Branch Proposal on the Manual Inputs tab first.'
      );
      return;
    }

    const isProtected =
      Boolean(manualEditedSections[section]) ||
      (section === 'sec8SourceOfRepayment' && manualInputs.sourceOfRepayment.trim() !== '');
    if (isProtected && !forceOverwrite) {
      return;
    }

    setScrutinizingSection(section);
    setAiError(null);

    try {
      const branchProposalSource = await prepareBranchProposalSource(branchProposalFile);
      const context = buildScrutinyContext('ic');

      const response = await fetchWithSourceCache(
        '/api/gemini/scrutinize-section',
        { section, context },
        branchProposalSource
      );

      const data = await parseJsonResponse<{
        result?: Record<string, unknown>;
        insufficient?: boolean;
        insufficientReason?: string | null;
        sourceId?: string;
      }>(response, 'AI scrutiny failed');
      if (data.sourceId) {
        serverCachedSourceIdRef.current = data.sourceId;
      } else if (branchProposalSource.sourceId) {
        serverCachedSourceIdRef.current = branchProposalSource.sourceId;
      }

      applyScrutinizedSectionResult(section, data.result ?? {});

      const isInsufficient = Boolean(data.insufficient);
      const reason =
        data.insufficientReason?.trim() ||
        'Relevant details for this section could not be found in the uploaded Branch Proposal.';
      setBatchScrutiny((prev) => {
        if (prev.total === 0 && prev.failedSections.length === 0 && !isInsufficient) {
          return prev;
        }
        const nextFailed = prev.failedSections.filter((f) => f.section !== section);
        const nextCompleted = prev.completedSections.includes(section)
          ? prev.completedSections
          : [...prev.completedSections, section];
        const nextInsufficient = isInsufficient
          ? [
              ...prev.insufficientSections.filter((item) => item.section !== section),
              { section, reason },
            ]
          : prev.insufficientSections.filter((item) => item.section !== section);
        return {
          ...prev,
          failedSections: nextFailed,
          completedSections: nextCompleted,
          insufficientSections: nextInsufficient,
        };
      });
    } catch (err) {
      setAiError(err instanceof Error ? err.message : 'Failed to scrutinize section.');
    } finally {
      setScrutinizingSection(null);
    }
  };

  const handleRetrySection = async (section: ScrutinySectionKey) => {
    if (batchScrutiny.running) return;
    if (!branchProposalFile) {
      setAiError(
        'Branch Proposal is not uploaded yet. Please upload the Branch Proposal on the Manual Inputs tab first.'
      );
      return;
    }
    if (batchScrutiny.retryingSections.includes(section)) return;

    setAiError(null);
    setBatchScrutiny((prev) => ({
      ...prev,
      retryingSections: prev.retryingSections.includes(section)
        ? prev.retryingSections
        : [...prev.retryingSections, section],
    }));

    try {
      const branchProposalSource = await prepareBranchProposalSource(branchProposalFile);
      const context = buildScrutinyContext('ic');

      const response = await fetchWithSourceCache(
        '/api/gemini/scrutinize-section',
        { section, context },
        branchProposalSource
      );

      const data = await parseJsonResponse<{
        result?: Record<string, unknown>;
        insufficient?: boolean;
        insufficientReason?: string | null;
        sourceId?: string;
      }>(response, 'Retry failed');
      if (data.sourceId) {
        serverCachedSourceIdRef.current = data.sourceId;
      } else if (branchProposalSource.sourceId) {
        serverCachedSourceIdRef.current = branchProposalSource.sourceId;
      }

      applyScrutinizedSectionResult(section, data.result ?? {});

      const isInsufficient = Boolean(data.insufficient);
      const reason =
        data.insufficientReason?.trim() ||
        'Relevant information is unavailable or insufficient in the uploaded Branch Proposal.';

      setBatchScrutiny((prev) => ({
        ...prev,
        failedSections: prev.failedSections.filter((f) => f.section !== section),
        completedSections: prev.completedSections.includes(section)
          ? prev.completedSections
          : [...prev.completedSections, section],
        insufficientSections: isInsufficient
          ? [
              ...prev.insufficientSections.filter((item) => item.section !== section),
              { section, reason },
            ]
          : prev.insufficientSections.filter((item) => item.section !== section),
      }));
    } catch (err) {
      const errMsg = err instanceof Error ? err.message : 'Retry failed for this section.';
      setBatchScrutiny((prev) => {
        const others = prev.failedSections.filter((f) => f.section !== section);
        return {
          ...prev,
          completedSections: prev.completedSections.filter((s) => s !== section),
          failedSections: [...others, { section, error: errMsg }],
        };
      });
    } finally {
      setBatchScrutiny((prev) => ({
        ...prev,
        retryingSections: prev.retryingSections.filter((s) => s !== section),
      }));
    }
  };

  const handleScrutinizeAll = async () => {
    if (batchScrutiny.running || scrutinizingSection !== null) return;
    if (!branchProposalFile) {
      setAiError(
        'Branch Proposal is not uploaded yet. Please upload the Branch Proposal on the Manual Inputs tab first.'
      );
      return;
    }

    batchAbortRef.current?.abort();
    const controller = new AbortController();
    batchAbortRef.current = controller;

    setAiError(null);

    const skippedManualSections: ScrutinySectionKey[] = [];
    const targetSections: ScrutinySectionKey[] = [];
    for (const s of ALL_SCRUTINY_SECTIONS) {
      const isProtected =
        Boolean(manualEditedSections[s]) ||
        (s === 'sec8SourceOfRepayment' && manualInputs.sourceOfRepayment.trim() !== '');
      if (isProtected) {
        skippedManualSections.push(s);
      } else {
        targetSections.push(s);
      }
    }

    if (targetSections.length === 0) {
      setBatchScrutiny({
        running: false,
        total: 0,
        targetSections: [],
        completedSections: [],
        retryingSections: [],
        failedSections: [],
        insufficientSections: [],
        skippedManualSections,
      });
      return;
    }

    setBatchScrutiny({
      running: true,
      total: targetSections.length,
      targetSections,
      completedSections: [],
      retryingSections: [],
      failedSections: [],
      insufficientSections: [],
      skippedManualSections,
    });

    try {
      const branchProposalSource = await prepareBranchProposalSource(branchProposalFile);
      const context = buildScrutinyContext('ic');

      const response = await fetchWithSourceCache(
        '/api/gemini/scrutinize-all',
        { sections: targetSections, context },
        branchProposalSource,
        controller.signal
      );

      await assertSseResponseOrThrow(response, 'Batch AI scrutiny failed');

      if (branchProposalSource.sourceId) {
        serverCachedSourceIdRef.current = branchProposalSource.sourceId;
      }

      const reader = response.body?.getReader();
      if (!reader) {
        throw new Error('Streaming response body unavailable from server.');
      }

      const decoder = new TextDecoder();
      let buffer = '';

      const processFrames = (frames: string[]) => {
        const sectionItems: Array<{ section: ScrutinySectionKey; result: Record<string, unknown> }> = [];
        const sectionStatuses: Array<{
          section: ScrutinySectionKey;
          insufficient: boolean;
          reason: string;
        }> = [];
        const errorItems: Array<{ section: ScrutinySectionKey; error: string }> = [];

        for (const frame of frames) {
          const dataLines = frame
            .split(/\r?\n/)
            .filter((line) => line.startsWith('data:'))
            .map((line) => line.slice(5).trimStart());
          if (dataLines.length === 0) continue;

          const rawPayload = dataLines.join('\n');
          try {
            const event = JSON.parse(rawPayload) as {
              type: 'section' | 'error' | 'fatal' | 'done';
              section?: ScrutinySectionKey;
              result?: Record<string, unknown>;
              insufficient?: boolean;
              insufficientReason?: string | null;
              error?: string;
              sourceId?: string;
            };

            if (event.type === 'section' && event.section && event.result) {
              const sec = event.section;
              sectionItems.push({ section: sec, result: event.result });
              sectionStatuses.push({
                section: sec,
                insufficient: Boolean(event.insufficient),
                reason:
                  event.insufficientReason?.trim() ||
                  'Relevant details for this section were not available or sufficient in the uploaded Branch Proposal.',
              });
            } else if (event.type === 'error' && event.section) {
              errorItems.push({
                section: event.section,
                error: event.error || 'Failed to generate section.',
              });
            } else if (event.type === 'fatal') {
              setAiError(event.error || 'Batch AI scrutiny failed.');
            } else if (event.type === 'done' && event.sourceId) {
              serverCachedSourceIdRef.current = event.sourceId;
            }
          } catch {
            // Ignore malformed SSE chunk
          }
        }

        if (sectionItems.length > 0) {
          applyScrutinizedSectionsBatch(sectionItems, false);
        }

        if (sectionStatuses.length > 0 || errorItems.length > 0) {
          setBatchScrutiny((prev) => {
            let nextCompleted = prev.completedSections;
            let nextFailed = prev.failedSections;
            let nextInsufficient = prev.insufficientSections;

            for (const st of sectionStatuses) {
              if (!nextCompleted.includes(st.section)) {
                nextCompleted =
                  nextCompleted === prev.completedSections
                    ? [...nextCompleted, st.section]
                    : (nextCompleted.push(st.section), nextCompleted);
              }
              nextFailed = nextFailed.filter((f) => f.section !== st.section);
              nextInsufficient = st.insufficient
                ? [
                    ...nextInsufficient.filter((item) => item.section !== st.section),
                    { section: st.section, reason: st.reason },
                  ]
                : nextInsufficient.filter((item) => item.section !== st.section);
            }

            for (const errItem of errorItems) {
              nextFailed = [
                ...nextFailed.filter((f) => f.section !== errItem.section),
                { section: errItem.section, error: errItem.error },
              ];
            }

            return {
              ...prev,
              completedSections: nextCompleted,
              failedSections: nextFailed,
              insufficientSections: nextInsufficient,
            };
          });
        }
      };

      while (true) {
        const { done, value } = await reader.read();
        if (done) {
          const tail = buffer + decoder.decode();
          if (tail.trim()) {
            processFrames(tail.split(/\r?\n\r?\n/));
          }
          break;
        }
        buffer += decoder.decode(value, { stream: true });

        const frames = buffer.split(/\r?\n\r?\n/);
        buffer = frames.pop() ?? '';
        if (frames.length > 0) {
          processFrames(frames);
        }
      }
    } catch (err) {
      if ((err as { name?: string })?.name !== 'AbortError') {
        setAiError(err instanceof Error ? err.message : 'Batch AI scrutiny failed.');
      }
    } finally {
      void persistCentralCaseNow(true);
      setBatchScrutiny((prev) => {
        const completedSet = new Set(prev.completedSections);
        const failedSet = new Set(prev.failedSections.map((f) => f.section));
        const uncompleted = prev.targetSections.filter(
          (sec) => !completedSet.has(sec) && !failedSet.has(sec)
        );
        const extraFailures = uncompleted.map((sec) => ({
          section: sec,
          error: 'Section did not finish generating — click Retry to scrutinize this section.',
        }));
        return {
          ...prev,
          running: false,
          failedSections: [...prev.failedSections, ...extraFailures],
        };
      });
      if (batchAbortRef.current === controller) {
        batchAbortRef.current = null;
      }
    }
  };

  const handleUpdateAiSections = (
    updater: (prev: AiSectionsData) => AiSectionsData,
    markManualSection?: ScrutinySectionKey
  ) => {
    setAiSectionsSync(updater);
    if (markManualSection) {
      setManualEditedSectionsSync((prev) => ({ ...prev, [markManualSection]: true }));
      clearOverridesForSection(markManualSection);
    }
    syncAppliedSnapshotFromRef();
    void persistCentralCaseNow(false);
  };

  const handleStageInputChange = (field: keyof StageMemoInputs, value: string) => {
    setStageInputsSync((prev) => ({ ...prev, [field]: value }));
    syncAppliedSnapshotFromRef();
    void persistCentralCaseNow(false);
  };

  const handleUpdateStageAiSection = (
    section: StageScrutinySectionKey,
    value: string,
    markManual = true
  ) => {
    setStageAiSectionsSync((prev) => ({ ...prev, [section]: value }));
    if (markManual) {
      setStageManualEditedSectionsSync((prev) => ({ ...prev, [section]: true }));
    }
    syncAppliedSnapshotFromRef();
    void persistCentralCaseNow(false);
  };

  const setStageSourceUploadsSync = (
    updater:
      | Record<StageMemoDocType, StageSourceUploadState>
      | ((
          prev: Record<StageMemoDocType, StageSourceUploadState>
        ) => Record<StageMemoDocType, StageSourceUploadState>)
  ) => {
    const next =
      typeof updater === 'function' ? updater(caseRef.current.stageSourceUploads) : updater;
    caseRef.current = { ...caseRef.current, stageSourceUploads: next };
    setStageSourceUploads(next);
  };

  const handleClearStageSource = (docType: StageMemoDocType) => {
    delete stageSourceFilesRef.current[docType];
    delete stageServerCachedSourceIdsRef.current[docType];
    const defaults = createInitialStageSourceUploads();
    setStageSourceUploadsSync((prev) => ({
      ...prev,
      [docType]: defaults[docType],
    }));
    void persistCentralCaseNow(true);
  };

  const handleClearAllData = async () => {
    setClearingCaseData(true);
    try {
      batchAbortRef.current?.abort();
      stageBatchAbortRef.current?.abort();
      batchAbortRef.current = null;
      stageBatchAbortRef.current = null;

      if (persistTimerRef.current !== null) {
        window.clearTimeout(persistTimerRef.current);
        persistTimerRef.current = null;
      }

      serverCachedSourceIdRef.current = null;
      stageSourceFilesRef.current = {};
      stageServerCachedSourceIdsRef.current = {};
      cacheWarmPromiseRef.current = null;
      caseDirtyRef.current = false;

      const freshStageUploads = createInitialStageSourceUploads();
      const emptyCase = {
        extractedDocs: INITIAL_EXTRACTED_DOCS,
        manualInputs: INITIAL_MANUAL_INPUTS,
        aiSections: INITIAL_AI_SECTIONS,
        overrides: {},
        manualEditedSections: {},
        stageInputs: INITIAL_STAGE_INPUTS,
        stageAiSections: INITIAL_STAGE_AI_SECTIONS,
        stageManualEditedSections: {},
        stageSourceUploads: freshStageUploads,
      };
      caseRef.current = emptyCase;

      setDocMetasSync(INITIAL_DOC_METAS);
      setExtractedDocs(INITIAL_EXTRACTED_DOCS);
      setManualInputs(INITIAL_MANUAL_INPUTS);
      setAiSections(INITIAL_AI_SECTIONS);
      setOverrides({});
      setManualEditedSections({});
      setStageInputs(INITIAL_STAGE_INPUTS);
      setStageAiSections(INITIAL_STAGE_AI_SECTIONS);
      setStageManualEditedSections({});
      setStageSourceUploads(freshStageUploads);
      setBranchProposalFile(null);
      setScrutinizingSection(null);
      setScrutinizingStageSection(null);
      setBatchScrutiny(INITIAL_BATCH_SCRUTINY);
      setStageBatchScrutiny(INITIAL_STAGE_BATCH_SCRUTINY);
      setAiError(null);

      await clearCentralCaseState();

      const emptyMemo = resolveICMemo(
        INITIAL_EXTRACTED_DOCS,
        INITIAL_MANUAL_INPUTS,
        INITIAL_AI_SECTIONS,
        {},
        {}
      );
      setAppliedSnapshot({
        appliedAt: null,
        version: 0,
        memo: emptyMemo,
        stageInputs: INITIAL_STAGE_INPUTS,
        stageAiSections: INITIAL_STAGE_AI_SECTIONS,
      });
      setHasPendingChanges(false);
      setApplySignal((s) => s + 1);
      setApplyNotice(
        'All case data has been cleared. Default document templates and software settings remain preserved.'
      );
      setShowClearConfirm(false);
    } finally {
      setClearingCaseData(false);
    }
  };

  const handleUploadStageSource = async (docType: StageMemoDocType, file: File) => {
    stageSourceFilesRef.current[docType] = file;
    setAiError(null);
    setStageSourceUploadsSync((prev) => ({
      ...prev,
      [docType]: {
        ...prev[docType],
        uploaded: false,
        fileName: file.name,
        processing: true,
        stage: 'Preparing document for AI reading & scrutiny...',
        error: null,
      },
    }));

    try {
      const prepared = await prepareMemoDocumentSource(file, (stageMsg) => {
        setStageSourceUploadsSync((prev) => ({
          ...prev,
          [docType]: {
            ...prev[docType],
            stage: stageMsg,
          },
        }));
      });

      setStageSourceUploadsSync((prev) => ({
        ...prev,
        [docType]: {
          ...prev[docType],
          stage: 'Running Gemini AI extraction, cross-verification & formatting...',
        },
      }));

      const cachedId = prepared.sourceId;
      const existingSnapshot = caseRef.current;
      const currentResolved = resolveICMemo(
        existingSnapshot.extractedDocs,
        existingSnapshot.manualInputs,
        existingSnapshot.aiSections,
        existingSnapshot.overrides,
        existingSnapshot.manualEditedSections
      );

      const response = await fetch('/api/gemini/extract-stage-source', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          targetDocType: docType,
          sourceId: cachedId,
          memoSource: prepared,
          stageInputs: existingSnapshot.stageInputs,
          existingContext: { placeholders: currentResolved.placeholders },
        }),
      });

      const data = await parseJsonResponse<{
        sourceId?: string;
        fileName?: string;
        extracted?: {
          placeholders?: Array<{ key: string; value: string }>;
          manualInputs?: Partial<ManualInputsData>;
          extractedStageInputs?: Partial<StageMemoInputs>;
          narratives?: {
            justification?: string;
            sec1NatureOfBusiness?: string;
            sec8SourceOfRepayment?: string;
            clientPrayer?: string;
            branchVisit?: string;
            sec13BranchRecommendation?: string;
            sec14IccdRecommendation?: string;
            samdObservation?: string;
            carriedMecIcRecommendation?: string;
            carriedBoardMecRecommendation?: string;
          };
          structuredTables?: {
            sec2Limit?: string;
            sec2Validity?: string;
            sec2Remarks?: string;
            sec2NonFundedNote?: string;
            section3Rows?: BranchProposalData['section3Rows'];
            sec4CibDate?: string;
            section4Rows?: BranchProposalData['section4Rows'];
            sec5StatusNote?: string;
            section5Rows?: BranchProposalData['section5Rows'];
            sec6StatusNote?: string;
            section6Rows?: BranchProposalData['section6Rows'];
            sec7PrimaryNote?: string;
            sec7PrimaryTotal?: string;
            sec7CollateralNote?: string;
            sec7CollateralSchedules?: string[];
            sec7CoveragePct?: string;
            sec7OthersNote?: string;
            section9Rows?: BranchProposalData['section9Rows'];
            sec10ParticularsPairs?: Array<{ key: string; value: string }>;
            section10Directors?: BranchProposalData['section10Directors'];
            sec11FinancialsNote?: string;
            section11Financials?: BranchProposalData['section11Rows'];
            previousBoardDecisions?: string;
          };
          targetStageSections?: Partial<Record<StageScrutinySectionKey, string>>;
          auditFindings?: StageSourceAuditResult;
        };
      }>(response, 'Stage memo AI extraction failed');

      if (data.sourceId) {
        stageServerCachedSourceIdsRef.current[docType] = data.sourceId;
      }

      const ext = data.extracted ?? {};
      let countExtracted = 0;

      // 1. Placeholders -> overrides (so uploaded finalized memo values preserve 100% accuracy across all tabs)
      const nextOverrides: Record<string, string> = { ...caseRef.current.overrides };
      for (const item of ext.placeholders || []) {
        const code = (item?.key || '').trim();
        const val = (item?.value || '').trim();
        if (code && val) {
          nextOverrides[code] = val;
          countExtracted++;
        }
      }
      setOverridesSync(nextOverrides);

      // 2. Manual Inputs
      if (ext.manualInputs) {
        setManualInputsSync((prev) => {
          const next = { ...prev };
          (Object.keys(ext.manualInputs || {}) as Array<keyof ManualInputsData>).forEach((k) => {
            const v = (ext.manualInputs?.[k] ?? '').trim();
            if (v) {
              next[k] = v;
              countExtracted++;
            }
          });
          return next;
        });
      }

      // 3. Stage Inputs (fill empty stage order/date fields if extracted from uploaded doc)
      if (ext.extractedStageInputs) {
        setStageInputsSync((prev) => {
          const next = { ...prev };
          (Object.keys(ext.extractedStageInputs || {}) as Array<keyof StageMemoInputs>).forEach((k) => {
            const v = (ext.extractedStageInputs?.[k] ?? '').trim();
            if (v && !next[k].trim()) {
              next[k] = v;
              countExtracted++;
            }
          });
          return next;
        });
      }

      // 4. Structured Tables -> BranchProposalData
      const st = ext.structuredTables;
      if (st) {
        const sec10Map: Record<string, string> = {};
        for (const kv of st.sec10ParticularsPairs || []) {
          if (kv?.key && kv?.value && kv.value.trim()) {
            sec10Map[kv.key.trim()] = kv.value.trim();
            countExtracted++;
          }
        }
        const emptyParticulars = {
          nameOfClient: '',
          uniqueCustomerId: '',
          awcdMsdAccount: '',
          businessTypeNature: '',
          businessManagementType: '',
          clientType: '',
          keyPersonPhone: '',
          address: '',
          officeBusinessAddress: '',
          factoryAddress: '',
          purposeOfInvestment: '',
          limitOfInvestment: '',
          modeOfInvestment: '',
          lastVisitDate: '',
          stockPosition: '',
          fixedAsset: '',
          businessNetWorth: '',
          insuranceCoverage: '',
          tradeLicenseValidity: '',
          otherRegulatoryDocuments: '',
          dateOfOriginalInvestment: '',
          firstDisbursementManager: '',
          firstDateOfClassification: '',
        };
        setExtractedDocsSync((prev) => {
          const curBp = prev.branchProposal;
          const nextBp: BranchProposalData = {
            referenceNo: nextOverrides['#64#'] || curBp?.referenceNo || '',
            proposalDate:
              nextOverrides['#1#'] || ext.manualInputs?.proposalDate || curBp?.proposalDate || '',
            branchName: nextOverrides['#3#'] || curBp?.branchName || '',
            accountNo: nextOverrides['#2#'] || curBp?.accountNo || '',
            subjectText: curBp?.subjectText || '',
            limitOriginal:
              sec10Map.limitOfInvestment || nextOverrides['#6#'] || curBp?.limitOriginal || '',
            limitRSD: curBp?.limitRSD || '',
            securityMV: nextOverrides['#33#'] || curBp?.securityMV || '',
            securityFSV: nextOverrides['#34#'] || curBp?.securityFSV || '',
            justificationText:
              ext.narratives?.justification ||
              nextOverrides['#51#'] ||
              curBp?.justificationText ||
              '',
            previousBoardDecisions:
              st.previousBoardDecisions || curBp?.previousBoardDecisions || 'Nil',
            natureOfBusiness:
              ext.narratives?.sec1NatureOfBusiness ||
              nextOverrides['#54#'] ||
              curBp?.natureOfBusiness ||
              '',
            sec2Limit: st.sec2Limit || curBp?.sec2Limit || '',
            sec2Validity: st.sec2Validity || curBp?.sec2Validity || '',
            sec2Overdue: nextOverrides['#59#'] || curBp?.sec2Overdue || '',
            sec2Recovery: nextOverrides['#60#'] || curBp?.sec2Recovery || '',
            sec2Remarks:
              [st.sec2Remarks, st.sec2NonFundedNote].filter(Boolean).join(' ') ||
              curBp?.sec2Remarks ||
              '',
            section3AsOnDate: nextOverrides['#68#'] || curBp?.section3AsOnDate || '',
            section3Rows:
              st.section3Rows && st.section3Rows.length > 0
                ? st.section3Rows
                : curBp?.section3Rows || [],
            section4CibDate: st.sec4CibDate || curBp?.section4CibDate || '',
            section4Rows:
              st.section4Rows && st.section4Rows.length > 0
                ? st.section4Rows
                : curBp?.section4Rows || [],
            section5StatusNote: st.sec5StatusNote || curBp?.section5StatusNote || '',
            section5Rows:
              st.section5Rows && st.section5Rows.length > 0
                ? st.section5Rows
                : curBp?.section5Rows || [],
            section6StatusNote: st.sec6StatusNote || curBp?.section6StatusNote || '',
            section6Rows:
              st.section6Rows && st.section6Rows.length > 0
                ? st.section6Rows
                : curBp?.section6Rows || [],
            section7PrimaryNote: st.sec7PrimaryNote || curBp?.section7PrimaryNote || '',
            section7PrimaryRows: curBp?.section7PrimaryRows || [],
            section7PrimaryTotal: st.sec7PrimaryTotal || curBp?.section7PrimaryTotal || '',
            section7CollateralNote: st.sec7CollateralNote || curBp?.section7CollateralNote || '',
            section7CollateralSchedules:
              st.sec7CollateralSchedules && st.sec7CollateralSchedules.length > 0
                ? st.sec7CollateralSchedules
                : curBp?.section7CollateralSchedules || [],
            section7ValuationRows: curBp?.section7ValuationRows || [],
            section7CoveragePct: st.sec7CoveragePct || curBp?.section7CoveragePct || '',
            section7Others: st.sec7OthersNote || curBp?.section7Others || '',
            section8SourceOfRepayment:
              ext.narratives?.sec8SourceOfRepayment ||
              ext.manualInputs?.sourceOfRepayment ||
              curBp?.section8SourceOfRepayment ||
              '',
            section9Rows:
              st.section9Rows && st.section9Rows.length > 0
                ? st.section9Rows
                : curBp?.section9Rows || [],
            section10Particulars: {
              ...emptyParticulars,
              ...(curBp?.section10Particulars || {}),
              ...(sec10Map as Partial<typeof emptyParticulars>),
            },
            section10Directors:
              st.section10Directors && st.section10Directors.length > 0
                ? st.section10Directors
                : curBp?.section10Directors || [],
            section10BankDirectors: curBp?.section10BankDirectors || [],
            section11FinancialsNote: st.sec11FinancialsNote || curBp?.section11FinancialsNote || '',
            section11Rows:
              st.section11Financials && st.section11Financials.length > 0
                ? st.section11Financials
                : curBp?.section11Rows || [],
            section12BranchVisitReport:
              ext.narratives?.branchVisit || curBp?.section12BranchVisitReport || '',
            section15BranchBicRec: curBp?.section15BranchBicRec || '',
            section16BranchRec:
              ext.narratives?.sec13BranchRecommendation || curBp?.section16BranchRec || '',
          };
          return { ...prev, branchProposal: nextBp };
        });
      }

      // 5. Carried-forward IC Narratives -> aiSections (respecting manual edits)
      if (ext.narratives) {
        const nar = ext.narratives;
        const edited = caseRef.current.manualEditedSections;
        setAiSectionsSync((prev) => ({
          ...prev,
          justification:
            !edited.justification && nar.justification?.trim()
              ? nar.justification.trim()
              : prev.justification,
          sec1NatureOfBusiness:
            !edited.sec1NatureOfBusiness && nar.sec1NatureOfBusiness?.trim()
              ? nar.sec1NatureOfBusiness.trim()
              : prev.sec1NatureOfBusiness,
          sec8SourceOfRepayment:
            !edited.sec8SourceOfRepayment && nar.sec8SourceOfRepayment?.trim()
              ? nar.sec8SourceOfRepayment.trim()
              : prev.sec8SourceOfRepayment,
          clientPrayer:
            !edited.clientPrayer && nar.clientPrayer?.trim()
              ? nar.clientPrayer.trim()
              : prev.clientPrayer,
          branchVisit:
            !edited.branchVisit && nar.branchVisit?.trim()
              ? nar.branchVisit.trim()
              : prev.branchVisit,
          sec13BranchRecommendation:
            !edited.sec13BranchRecommendation && nar.sec13BranchRecommendation?.trim()
              ? nar.sec13BranchRecommendation.trim()
              : prev.sec13BranchRecommendation,
          sec14IccdRecommendation:
            !edited.sec14IccdRecommendation && nar.sec14IccdRecommendation?.trim()
              ? nar.sec14IccdRecommendation.trim()
              : prev.sec14IccdRecommendation,
          samdObservation:
            !edited.samdObservation && nar.samdObservation?.trim()
              ? nar.samdObservation.trim()
              : prev.samdObservation,
        }));
      }

      // 6. Target Stage AI Sections (never overwrite manually edited sections; Section 16 & 17 remain auto-synced to user inputs unless manually edited)
      const stageEdits = caseRef.current.stageManualEditedSections;
      const ts = ext.targetStageSections || {};
      setStageAiSectionsSync((prev) => {
        const next = { ...prev };
        for (const k of [
          'sanctionCrossVerifiedSummary',
        ] as StageScrutinySectionKey[]) {
          const val = (ts[k] ?? '').trim();
          if (val && !stageEdits[k]) {
            next[k] = val;
            countExtracted++;
          }
        }
        return next;
      });

      const audit: StageSourceAuditResult = {
        summary:
          ext.auditFindings?.summary?.trim() ||
          `Extracted and carried forward structured data from "${file.name}".`,
        verifiedPoints: Array.isArray(ext.auditFindings?.verifiedPoints)
          ? ext.auditFindings.verifiedPoints.filter((s) => typeof s === 'string' && s.trim())
          : [],
        discrepanciesOrMissing: Array.isArray(ext.auditFindings?.discrepanciesOrMissing)
          ? ext.auditFindings.discrepanciesOrMissing.filter((s) => typeof s === 'string' && s.trim())
          : [],
      };

      const extractedTime = new Date().toLocaleTimeString([], {
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
      });

      setStageSourceUploadsSync((prev) => ({
        ...prev,
        [docType]: {
          ...prev[docType],
          uploaded: true,
          fileName: file.name,
          processing: false,
          stage: null,
          error: null,
          extractedAt: extractedTime,
          extractedFieldCount: countExtracted,
          audit,
        },
      }));

      syncAppliedSnapshotFromRef();
      await persistCentralCaseNow(true);
    } catch (err) {
      const msg =
        err instanceof Error
          ? err.message
          : 'Failed to read and extract the uploaded memo document.';
      setStageSourceUploadsSync((prev) => ({
        ...prev,
        [docType]: {
          ...prev[docType],
          processing: false,
          stage: null,
          error: msg,
        },
      }));
    }
  };

  const fetchStageEndpoint = async (
    url: string,
    payload: Record<string, unknown>,
    signal?: AbortSignal
  ): Promise<Response> => {
    const docType = payload.docType as StageMemoDocType | undefined;
    const stageFile = docType ? stageSourceFilesRef.current[docType] : undefined;
    if (stageFile && docType) {
      const stageSource = await prepareMemoDocumentSource(stageFile);
      return fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        signal,
        body: JSON.stringify({
          ...payload,
          sourceId: stageServerCachedSourceIdsRef.current[docType] || stageSource.sourceId,
          stageSource,
        }),
      });
    }
    if (branchProposalFile) {
      const source = await prepareBranchProposalSource(branchProposalFile);
      return fetchWithSourceCache(url, payload, source, signal);
    }
    return fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      signal,
      body: JSON.stringify(payload),
    });
  };

  const hasStageOrCaseSource = (docType: StageMemoDocType) =>
    Boolean(
      branchProposalFile ||
        caseRef.current.extractedDocs.branchProposal ||
        resolvedICMemo.placeholders['#2#'] ||
        stageSourceFilesRef.current[docType] ||
        caseRef.current.stageSourceUploads[docType]?.uploaded
    );

  const handleScrutinizeStageSection = async (
    docType: StageMemoDocType,
    section: StageScrutinySectionKey,
    forceOverwrite = false
  ) => {
    if (stageBatchScrutiny.running) return;
    if (!hasStageOrCaseSource(docType)) {
      setAiError(
        'Upload the case documents on Manual Inputs OR upload the completed previous-stage memo right above so AI can scrutinize and populate this stage.'
      );
      return;
    }
    if (stageManualEditedSections[section] && !forceOverwrite) {
      return;
    }

    setScrutinizingStageSection(section);
    setAiError(null);

    try {
      const context = buildScrutinyContext();
      const response = await fetchStageEndpoint('/api/gemini/scrutinize-stage-section', {
        docType,
        section,
        context,
      });

      const data = await parseJsonResponse<{
        result?: { text?: string };
        insufficient?: boolean;
        insufficientReason?: string | null;
        sourceId?: string;
      }>(response, 'Stage AI scrutiny failed');
      if (data.sourceId) {
        serverCachedSourceIdRef.current = data.sourceId;
      }

      const text = String(data.result?.text ?? '').trim();
      setStageAiSectionsSync((prev) => ({ ...prev, [section]: text }));
      setStageManualEditedSectionsSync((prev) => ({ ...prev, [section]: false }));
      syncAppliedSnapshotFromRef();
      void persistCentralCaseNow(true);

      const isInsufficient = Boolean(data.insufficient);
      const reason =
        data.insufficientReason?.trim() ||
        'Some reference details were incomplete in the current case documents.';
      setStageBatchScrutiny((prev) => {
        if (prev.docType !== docType && prev.total === 0) return prev;
        return {
          ...prev,
          docType,
          failedSections: prev.failedSections.filter((f) => f.section !== section),
          completedSections: prev.completedSections.includes(section)
            ? prev.completedSections
            : [...prev.completedSections, section],
          insufficientSections: isInsufficient
            ? [
                ...prev.insufficientSections.filter((item) => item.section !== section),
                { section, reason },
              ]
            : prev.insufficientSections.filter((item) => item.section !== section),
        };
      });
    } catch (err) {
      setAiError(err instanceof Error ? err.message : 'Failed to scrutinize stage section.');
    } finally {
      setScrutinizingStageSection(null);
    }
  };

  const handleRetryStageSection = async (
    docType: StageMemoDocType,
    section: StageScrutinySectionKey
  ) => {
    if (stageBatchScrutiny.running) return;
    if (stageBatchScrutiny.retryingSections.includes(section)) return;

    setAiError(null);
    setStageBatchScrutiny((prev) => ({
      ...prev,
      docType,
      retryingSections: prev.retryingSections.includes(section)
        ? prev.retryingSections
        : [...prev.retryingSections, section],
    }));

    try {
      const context = buildScrutinyContext();
      const response = await fetchStageEndpoint('/api/gemini/scrutinize-stage-section', {
        docType,
        section,
        context,
      });

      const data = await parseJsonResponse<{
        result?: { text?: string };
        insufficient?: boolean;
        insufficientReason?: string | null;
        sourceId?: string;
      }>(response, 'Retry failed');
      if (data.sourceId) {
        serverCachedSourceIdRef.current = data.sourceId;
      }

      const text = String(data.result?.text ?? '').trim();
      setStageAiSectionsSync((prev) => ({ ...prev, [section]: text }));
      setStageManualEditedSectionsSync((prev) => ({ ...prev, [section]: false }));
      syncAppliedSnapshotFromRef();
      void persistCentralCaseNow(true);

      const isInsufficient = Boolean(data.insufficient);
      const reason =
        data.insufficientReason?.trim() ||
        'Some reference details were incomplete in the current case documents.';

      setStageBatchScrutiny((prev) => ({
        ...prev,
        docType,
        failedSections: prev.failedSections.filter((f) => f.section !== section),
        completedSections: prev.completedSections.includes(section)
          ? prev.completedSections
          : [...prev.completedSections, section],
        insufficientSections: isInsufficient
          ? [
              ...prev.insufficientSections.filter((item) => item.section !== section),
              { section, reason },
            ]
          : prev.insufficientSections.filter((item) => item.section !== section),
      }));
    } catch (err) {
      const errMsg = err instanceof Error ? err.message : 'Retry failed for this section.';
      setStageBatchScrutiny((prev) => ({
        ...prev,
        docType,
        completedSections: prev.completedSections.filter((s) => s !== section),
        failedSections: [
          ...prev.failedSections.filter((f) => f.section !== section),
          { section, error: errMsg },
        ],
      }));
    } finally {
      setStageBatchScrutiny((prev) => ({
        ...prev,
        retryingSections: prev.retryingSections.filter((s) => s !== section),
      }));
    }
  };

  const handleScrutinizeStageAll = async (docType: StageMemoDocType) => {
    if (stageBatchScrutiny.running || scrutinizingStageSection !== null) return;
    if (!hasStageOrCaseSource(docType)) {
      setAiError(
        'Upload the case documents on Manual Inputs OR upload the completed previous-stage memo right above so AI can scrutinize and populate this stage.'
      );
      return;
    }

    stageBatchAbortRef.current?.abort();
    const controller = new AbortController();
    stageBatchAbortRef.current = controller;
    setAiError(null);

    const allStageKeys = STAGE_SECTIONS_BY_DOC[docType];
    const skippedManualSections: StageScrutinySectionKey[] = [];
    const targetSections: StageScrutinySectionKey[] = [];
    for (const s of allStageKeys) {
      if (stageManualEditedSections[s]) {
        skippedManualSections.push(s);
      } else {
        targetSections.push(s);
      }
    }

    if (targetSections.length === 0) {
      setStageBatchScrutiny({
        running: false,
        docType,
        total: allStageKeys.length,
        targetSections: [],
        completedSections: [],
        retryingSections: [],
        failedSections: [],
        insufficientSections: [],
        skippedManualSections,
      });
      return;
    }

    setStageBatchScrutiny({
      running: true,
      docType,
      total: targetSections.length,
      targetSections,
      completedSections: [],
      retryingSections: [],
      failedSections: [],
      insufficientSections: [],
      skippedManualSections,
    });

    try {
      const context = buildScrutinyContext('stage');
      const response = await fetchStageEndpoint(
        '/api/gemini/scrutinize-stage-all',
        { docType, sections: targetSections, context },
        controller.signal
      );

      await assertSseResponseOrThrow(response, 'Stage batch scrutiny failed');

      const reader = response.body?.getReader();
      if (!reader) {
        throw new Error('Streaming response body unavailable from server.');
      }

      const decoder = new TextDecoder();
      let buffer = '';

      const processStageFrames = (frames: string[]) => {
        const sectionItems: Array<{
          section: StageScrutinySectionKey;
          text: string;
          insufficient: boolean;
          reason: string;
        }> = [];
        const errorItems: Array<{ section: StageScrutinySectionKey; error: string }> = [];

        for (const frame of frames) {
          const dataLines = frame
            .split(/\r?\n/)
            .filter((line) => line.startsWith('data:'))
            .map((line) => line.slice(5).trimStart());
          if (dataLines.length === 0) continue;

          try {
            const event = JSON.parse(dataLines.join('\n')) as {
              type: 'section' | 'error' | 'fatal' | 'done';
              section?: StageScrutinySectionKey;
              result?: { text?: string };
              insufficient?: boolean;
              insufficientReason?: string | null;
              error?: string;
              sourceId?: string;
            };

            if (event.type === 'section' && event.section && event.result) {
              sectionItems.push({
                section: event.section,
                text: String(event.result.text ?? '').trim(),
                insufficient: Boolean(event.insufficient),
                reason:
                  event.insufficientReason?.trim() ||
                  'Some reference details were incomplete in the current case documents.',
              });
            } else if (event.type === 'error' && event.section) {
              errorItems.push({
                section: event.section,
                error: event.error || 'Failed to generate section.',
              });
            } else if (event.type === 'fatal') {
              setAiError(event.error || 'Stage batch AI scrutiny failed.');
            } else if (event.type === 'done' && event.sourceId) {
              serverCachedSourceIdRef.current = event.sourceId;
            }
          } catch {
            // Ignore malformed SSE chunk
          }
        }

        if (sectionItems.length > 0) {
          setStageAiSectionsSync((prev) => {
            const next = { ...prev };
            for (const item of sectionItems) {
              next[item.section] = item.text;
            }
            return next;
          });
          setStageManualEditedSectionsSync((prev) => {
            const next = { ...prev };
            for (const item of sectionItems) {
              next[item.section] = false;
            }
            return next;
          });
          syncAppliedSnapshotFromRef();
          void persistCentralCaseNow(false);
        }

        if (sectionItems.length > 0 || errorItems.length > 0) {
          setStageBatchScrutiny((prev) => {
            let nextCompleted = prev.completedSections;
            let nextFailed = prev.failedSections;
            let nextInsufficient = prev.insufficientSections;

            for (const item of sectionItems) {
              if (!nextCompleted.includes(item.section)) {
                nextCompleted =
                  nextCompleted === prev.completedSections
                    ? [...nextCompleted, item.section]
                    : (nextCompleted.push(item.section), nextCompleted);
              }
              nextFailed = nextFailed.filter((f) => f.section !== item.section);
              nextInsufficient = item.insufficient
                ? [
                    ...nextInsufficient.filter((entry) => entry.section !== item.section),
                    { section: item.section, reason: item.reason },
                  ]
                : nextInsufficient.filter((entry) => entry.section !== item.section);
            }

            for (const errItem of errorItems) {
              nextFailed = [
                ...nextFailed.filter((f) => f.section !== errItem.section),
                { section: errItem.section, error: errItem.error },
              ];
            }

            return {
              ...prev,
              completedSections: nextCompleted,
              failedSections: nextFailed,
              insufficientSections: nextInsufficient,
            };
          });
        }
      };

      while (true) {
        const { done, value } = await reader.read();
        if (done) {
          const tail = buffer + decoder.decode();
          if (tail.trim()) {
            processStageFrames(tail.split(/\r?\n\r?\n/));
          }
          break;
        }
        buffer += decoder.decode(value, { stream: true });

        const frames = buffer.split(/\r?\n\r?\n/);
        buffer = frames.pop() ?? '';
        if (frames.length > 0) {
          processStageFrames(frames);
        }
      }
    } catch (err) {
      if ((err as { name?: string })?.name !== 'AbortError') {
        setAiError(err instanceof Error ? err.message : 'Stage batch AI scrutiny failed.');
      }
    } finally {
      void persistCentralCaseNow(true);
      setStageBatchScrutiny((prev) => {
        const completedSet = new Set(prev.completedSections);
        const failedSet = new Set(prev.failedSections.map((f) => f.section));
        const uncompleted = prev.targetSections.filter(
          (sec) => !completedSet.has(sec) && !failedSet.has(sec)
        );
        const extraFailures = uncompleted.map((sec) => ({
          section: sec,
          error: 'Section did not finish generating — click Retry to scrutinize this section.',
        }));
        return {
          ...prev,
          running: false,
          failedSections: [...prev.failedSections, ...extraFailures],
        };
      });
      if (stageBatchAbortRef.current === controller) {
        stageBatchAbortRef.current = null;
      }
    }
  };

  const handleUploadTemplate = async (docType: MemoDocType, file: File) => {
    const record = await inspectDocxTemplate(file, docType);
    setTemplates((prev) => ({ ...prev, [docType]: record }));
    await saveTemplateRecord(record);
    syncAppliedSnapshotFromRef();
  };

  const handleRemoveTemplate = async (docType: MemoDocType) => {
    setTemplates((prev) => ({ ...prev, [docType]: null }));
    await deleteTemplateRecord(docType);
    syncAppliedSnapshotFromRef();
  };

  const handleApplyAllChanges = () => {
    void persistCentralCaseNow(true);
    syncAppliedSnapshotFromRef();
    const nowStr = new Date().toLocaleTimeString([], {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
    });
    setApplySignal((s) => s + 1);
    setApplyNotice(
      `All changes applied at ${nowStr} — IC Memo, MEC Memo, Board Memo, and Sanction outputs synchronized.`
    );
  };

  const triggerBlobDownload = (blob: Blob, fileName: string) => {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  };

  const prepareVerifiedSynchronizedExport = async (docType: MemoDocType) => {
    const activeTemplate = templates[docType];

    // Flush any queued microtask/state updates and persist central case data immediately
    await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
    await persistCentralCaseNow(true);
    const syncedMemo = syncAppliedSnapshotFromRef();
    const latest = caseRef.current;
    const syncedStageMemo =
      docType === 'ic'
        ? null
        : resolveStageMemo(
            docType,
            syncedMemo,
            latest.stageInputs,
            latest.stageAiSections,
            latest.stageManualEditedSections
          );

    const docxBlob = await generateSynchronizedDocxBlob(
      docType,
      syncedMemo,
      activeTemplate,
      syncedStageMemo
    );
    const verification = await verifyPopulatedDocxBlob(
      docxBlob,
      syncedMemo,
      docType,
      activeTemplate,
      syncedStageMemo
    );
    if (!verification.verified) {
      throw new Error(
        `Download verification failed: populated content for ${verification.missingLabels.join(', ')} was not included in the generated document.`
      );
    }

    const acct = syncedMemo.placeholders['#2#'] || 'Draft';
    const safeName = acct.replace(/[^A-Za-z0-9_-]/g, '_');
    return { docxBlob, safeName };
  };

  const handleDownloadDocx = async (docType: MemoDocType = 'ic') => {
    setDownloading(true);
    setAiError(null);
    try {
      const { docxBlob, safeName } = await prepareVerifiedSynchronizedExport(docType);
      triggerBlobDownload(docxBlob, `${DOC_FILE_PREFIX[docType]}_${safeName}.docx`);
    } catch (err) {
      setAiError(err instanceof Error ? err.message : 'Failed to generate .docx document.');
    } finally {
      setDownloading(false);
    }
  };

  const handleDownloadPdf = async (docType: MemoDocType = 'ic') => {
    setDownloadingPdf(true);
    setAiError(null);
    try {
      const { docxBlob, safeName } = await prepareVerifiedSynchronizedExport(docType);
      const pdfBlob = await generatePdfFromDocxBlob(docxBlob);
      triggerBlobDownload(pdfBlob, `${DOC_FILE_PREFIX[docType]}_${safeName}.pdf`);
    } catch (err) {
      setAiError(err instanceof Error ? err.message : 'Failed to generate .pdf document.');
    } finally {
      setDownloadingPdf(false);
    }
  };

  if (!entered) {
    return <LandingScreen onEnter={() => setEntered(true)} />;
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col">
      {/* Three-Zone Top Bar */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-20">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 h-16 flex items-center justify-between gap-8">
          {/* Zone 1: Brand Title */}
          <button
            type="button"
            onClick={() => setEntered(false)}
            className="font-display text-lg font-bold tracking-tight text-slate-900 hover:text-emerald-800 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
          >
            Memo Engine by Tuhin
          </button>

          {/* Zone 2: 5 Required Tabs in Exact Order */}
          <nav
            aria-label="Memo Workspace Tabs"
            className="flex items-center gap-1 overflow-x-auto"
          >
            {TABS.map((tab) => {
              const active = activeTab === tab.id;
              return (
                <button
                  key={tab.id}
                  type="button"
                  onClick={() => setActiveTab(tab.id)}
                  className={`px-3.5 py-2 text-sm font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                    active
                      ? 'bg-emerald-800 text-white'
                      : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
                  }`}
                >
                  {tab.label}
                </button>
              );
            })}
          </nav>

          {/* Zone 3: Primary Actions */}
          <div className="flex items-center gap-2 shrink-0">
            <button
              type="button"
              onClick={() => setShowClearConfirm(true)}
              className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold text-red-700 bg-red-50 hover:bg-red-100 border border-red-200 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              title="Clear all case inputs, uploaded documents, and generated memos while preserving templates"
            >
              <Trash2 className="w-3.5 h-3.5" />
              <span>Clear All Data</span>
            </button>

            <button
              type="button"
              onClick={handleApplyAllChanges}
              className={`hidden sm:inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-lg border transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                hasPendingChanges
                  ? 'text-amber-950 bg-amber-100 hover:bg-amber-200 border-amber-300'
                  : 'text-slate-700 bg-slate-100 hover:bg-slate-200 border-slate-300'
              }`}
            >
              <RefreshCw className="w-3.5 h-3.5" />
              <span>Apply All Changes</span>
            </button>

            {activeTab !== 'manual' ? (
              <div className="hidden sm:flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => void handleDownloadDocx(activeTab as MemoDocType)}
                  disabled={downloading || downloadingPdf}
                  className="px-3.5 py-2 text-xs font-semibold text-white bg-emerald-800 hover:bg-emerald-900 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer disabled:opacity-60"
                >
                  {downloading ? 'Building...' : 'Download .docx'}
                </button>
                <button
                  type="button"
                  onClick={() => void handleDownloadPdf(activeTab as MemoDocType)}
                  disabled={downloading || downloadingPdf}
                  className="px-3.5 py-2 text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer disabled:opacity-60"
                >
                  {downloadingPdf ? 'Rendering PDF...' : 'Download .pdf'}
                </button>
              </div>
            ) : (
              <button
                type="button"
                onClick={() => setActiveTab('ic')}
                className="hidden sm:inline-flex px-4 py-2 text-xs font-semibold text-slate-900 bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              >
                IC Memo
              </button>
            )}
          </div>
        </div>

        {applyNotice && (
          <div className="bg-emerald-50 border-t border-emerald-200 px-4 sm:px-6 py-2">
            <div className="max-w-7xl mx-auto flex items-center gap-2 text-xs font-medium text-emerald-900">
              <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
              <span>{applyNotice}</span>
            </div>
          </div>
        )}

        {aiError && activeTab !== 'ic' && (
          <div className="bg-red-50 border-t border-red-200 px-4 sm:px-6 py-2">
            <div className="max-w-7xl mx-auto flex items-center justify-between gap-2 text-xs font-medium text-red-800">
              <span>{aiError}</span>
              <button
                type="button"
                onClick={() => setAiError(null)}
                className="underline cursor-pointer shrink-0"
              >
                Dismiss
              </button>
            </div>
          </div>
        )}
      </header>

      {/* Active Tab Content */}
      <main className="flex-1 px-4 sm:px-6 py-6">
        {activeTab === 'manual' && (
          <ManualInputsTab
            docMetas={docMetas}
            extractedDocs={extractedDocs}
            manualInputs={manualInputs}
            overrides={overrides}
            resolvedMemo={resolvedICMemo}
            onFileUpload={handleFileUpload}
            onBatchUpload={handleBatchUpload}
            onManualChange={handleManualChange}
            onOverrideChange={handleOverrideChange}
            onResetOverrides={handleResetOverrides}
            onLoadSampleDocuments={handleLoadSampleDocuments}
            loadingSamples={loadingSamples}
            batchProgress={batchProgress}
            onGoToICMemo={() => setActiveTab('ic')}
            templates={templates}
            onUploadTemplate={handleUploadTemplate}
            onRemoveTemplate={handleRemoveTemplate}
            onApplyAllChanges={handleApplyAllChanges}
            hasPendingChanges={hasPendingChanges}
            onClearAllData={() => setShowClearConfirm(true)}
          />
        )}

        {activeTab === 'ic' && (
          <ICMemoPreview
            memo={resolvedICMemo}
            aiSections={aiSections}
            overrides={overrides}
            onOverrideChange={handleOverrideChange}
            onScrutinizeSection={handleScrutinizeSection}
            onRetrySection={handleRetrySection}
            onScrutinizeAll={handleScrutinizeAll}
            onDismissBatchSummary={() => setBatchScrutiny(INITIAL_BATCH_SCRUTINY)}
            batchScrutiny={batchScrutiny}
            onUpdateAiSections={handleUpdateAiSections}
            manualEditedSections={manualEditedSections}
            sourceOfRepaymentManual={manualInputs.sourceOfRepayment}
            onManualSourceOfRepaymentChange={handleManualSourceOfRepaymentChange}
            scrutinizingSection={scrutinizingSection}
            hasBranchProposal={branchProposalFile !== null}
            aiError={aiError}
            onDownloadDocx={() => handleDownloadDocx('ic')}
            onDownloadPdf={() => handleDownloadPdf('ic')}
            downloading={downloading}
            downloadingPdf={downloadingPdf}
            templates={templates}
            onUploadTemplate={handleUploadTemplate}
            onRemoveTemplate={handleRemoveTemplate}
            onApplyAllChanges={handleApplyAllChanges}
            hasPendingChanges={hasPendingChanges}
            applySignal={applySignal}
          />
        )}

        {activeTab === 'mec' && (
          <StageMemoTab
            docType="mec"
            stageMemo={resolvedMecMemo}
            stageInputs={stageInputs}
            stageAiSections={stageAiSections}
            stageManualEditedSections={stageManualEditedSections}
            scrutinizingStageSection={scrutinizingStageSection}
            stageBatchScrutiny={stageBatchScrutiny}
            hasBranchProposalFile={branchProposalFile !== null}
            hasAnyCaseData={Boolean(
              resolvedICMemo.placeholders['#2#'] || resolvedICMemo.branchProposal
            )}
            stageSourceUpload={stageSourceUploads.mec}
            onUploadStageSource={handleUploadStageSource}
            onClearStageSource={handleClearStageSource}
            aiError={aiError}
            onStageInputChange={handleStageInputChange}
            onUpdateStageAiSection={handleUpdateStageAiSection}
            onScrutinizeStageSection={handleScrutinizeStageSection}
            onRetryStageSection={handleRetryStageSection}
            onScrutinizeStageAll={handleScrutinizeStageAll}
            onDismissStageBatchSummary={() =>
              setStageBatchScrutiny(INITIAL_STAGE_BATCH_SCRUTINY)
            }
            templates={templates}
            onUploadTemplate={handleUploadTemplate}
            onRemoveTemplate={handleRemoveTemplate}
            onDownloadDocx={() => handleDownloadDocx('mec')}
            onDownloadPdf={() => handleDownloadPdf('mec')}
            downloading={downloading}
            downloadingPdf={downloadingPdf}
            onApplyAllChanges={handleApplyAllChanges}
            hasPendingChanges={hasPendingChanges}
            lastAppliedAt={appliedSnapshot.appliedAt}
            applySignal={applySignal}
          />
        )}

        {activeTab === 'board' && (
          <StageMemoTab
            docType="board"
            stageMemo={resolvedBoardMemo}
            stageInputs={stageInputs}
            stageAiSections={stageAiSections}
            stageManualEditedSections={stageManualEditedSections}
            scrutinizingStageSection={scrutinizingStageSection}
            stageBatchScrutiny={stageBatchScrutiny}
            hasBranchProposalFile={branchProposalFile !== null}
            hasAnyCaseData={Boolean(
              resolvedICMemo.placeholders['#2#'] || resolvedICMemo.branchProposal
            )}
            stageSourceUpload={stageSourceUploads.board}
            onUploadStageSource={handleUploadStageSource}
            onClearStageSource={handleClearStageSource}
            aiError={aiError}
            onStageInputChange={handleStageInputChange}
            onUpdateStageAiSection={handleUpdateStageAiSection}
            onScrutinizeStageSection={handleScrutinizeStageSection}
            onRetryStageSection={handleRetryStageSection}
            onScrutinizeStageAll={handleScrutinizeStageAll}
            onDismissStageBatchSummary={() =>
              setStageBatchScrutiny(INITIAL_STAGE_BATCH_SCRUTINY)
            }
            templates={templates}
            onUploadTemplate={handleUploadTemplate}
            onRemoveTemplate={handleRemoveTemplate}
            onDownloadDocx={() => handleDownloadDocx('board')}
            onDownloadPdf={() => handleDownloadPdf('board')}
            downloading={downloading}
            downloadingPdf={downloadingPdf}
            onApplyAllChanges={handleApplyAllChanges}
            hasPendingChanges={hasPendingChanges}
            lastAppliedAt={appliedSnapshot.appliedAt}
            applySignal={applySignal}
          />
        )}

        {activeTab === 'sanction' && (
          <StageMemoTab
            docType="sanction"
            stageMemo={resolvedSanctionMemo}
            stageInputs={stageInputs}
            stageAiSections={stageAiSections}
            stageManualEditedSections={stageManualEditedSections}
            scrutinizingStageSection={scrutinizingStageSection}
            stageBatchScrutiny={stageBatchScrutiny}
            hasBranchProposalFile={branchProposalFile !== null}
            hasAnyCaseData={Boolean(
              resolvedICMemo.placeholders['#2#'] || resolvedICMemo.branchProposal
            )}
            stageSourceUpload={stageSourceUploads.sanction}
            onUploadStageSource={handleUploadStageSource}
            onClearStageSource={handleClearStageSource}
            aiError={aiError}
            onStageInputChange={handleStageInputChange}
            onUpdateStageAiSection={handleUpdateStageAiSection}
            onScrutinizeStageSection={handleScrutinizeStageSection}
            onRetryStageSection={handleRetryStageSection}
            onScrutinizeStageAll={handleScrutinizeStageAll}
            onDismissStageBatchSummary={() =>
              setStageBatchScrutiny(INITIAL_STAGE_BATCH_SCRUTINY)
            }
            templates={templates}
            onUploadTemplate={handleUploadTemplate}
            onRemoveTemplate={handleRemoveTemplate}
            onDownloadDocx={() => handleDownloadDocx('sanction')}
            onDownloadPdf={() => handleDownloadPdf('sanction')}
            downloading={downloading}
            downloadingPdf={downloadingPdf}
            onApplyAllChanges={handleApplyAllChanges}
            hasPendingChanges={hasPendingChanges}
            lastAppliedAt={appliedSnapshot.appliedAt}
            applySignal={applySignal}
          />
        )}
      </main>

      {/* Confirmation Modal for Clear All Data */}
      {showClearConfirm && (
        <div
          className="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-[1px] flex items-center justify-center p-4"
          role="dialog"
          aria-modal="true"
          aria-labelledby="clear-data-dialog-title"
        >
          <div className="bg-white border border-slate-200 rounded-xl shadow-xl max-w-md w-full p-6 space-y-4">
            <div className="flex items-start gap-3">
              <div className="w-10 h-10 rounded-lg bg-red-50 border border-red-200 flex items-center justify-center shrink-0">
                <AlertTriangle className="w-5 h-5 text-red-700" />
              </div>
              <div className="space-y-1">
                <h2
                  id="clear-data-dialog-title"
                  className="font-display text-base font-bold text-slate-900"
                >
                  Clear All Case Data?
                </h2>
                <p className="text-xs text-slate-600 leading-relaxed">
                  This will reset all entered fields, uploaded documents, extracted figures, AI-generated or manually edited contents across Manual Inputs, IC Memo, MEC Memo, Board Memo, and Sanction tabs so you can start a fresh case.
                </p>
                <p className="text-xs font-medium text-emerald-900 bg-emerald-50 border border-emerald-200 rounded-md px-2.5 py-1.5 mt-2">
                  Your uploaded custom .docx templates, default bank formats, and software settings will NOT be affected.
                </p>
              </div>
            </div>

            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setShowClearConfirm(false)}
                disabled={clearingCaseData}
                className="px-4 py-2 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded-lg transition-colors cursor-pointer disabled:opacity-60"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={() => void handleClearAllData()}
                disabled={clearingCaseData}
                className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-red-700 hover:bg-red-800 rounded-lg transition-colors cursor-pointer disabled:opacity-60"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>{clearingCaseData ? 'Clearing...' : 'Yes, Clear All Data'}</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
