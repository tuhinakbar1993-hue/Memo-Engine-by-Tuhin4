export type DocumentKey =
  | 'branchProposal'
  | 'outstanding'
  | 'partyPerformance'
  | 'overdue'
  | 'accountStatement'
  | 'cibReport'
  | 'costOfFund';

export type MemoTabKey = 'manual' | 'ic' | 'mec' | 'board' | 'sanction';

export interface DocumentUploadMeta {
  key: DocumentKey;
  label: string;
  acceptedFormats: string;
  uploaded: boolean;
  fileName: string;
  parsing: boolean;
  stage?: string | null;
  error: string | null;
  rawTextPreview?: string;
}

export interface Section3Row {
  sl: string;
  nameOfConcern: string;
  bankOrNbfi: string;
  natureOfFacility: string;
  limit: string;
  outstanding: string;
  overdueEol: string;
  expiryDate: string;
  clStatus: string;
}

export interface Section4Row {
  status: string;
  unitBorrower: string;
  unitGuarantor: string;
  otherBorrower: string;
  otherGuarantor: string;
  remarks: string;
}

export interface Section5Row {
  sl: string;
  typeOfCase: string;
  caseNo: string;
  filingDate: string;
  amount: string;
  previousDate: string;
  nextDate: string;
  presentStatus: string;
}

export interface Section6Row {
  sl: string;
  authority: string;
  lastInspectionDate: string;
  objectionStatus: string;
  objections: string;
  branchReply: string;
}

export interface Section7PrimaryRow {
  particulars: string;
  value: string;
}

export interface Section7ValuationRow {
  schedule: string;
  origBranchMv: string;
  origBranchFsv: string;
  origSurvMv: string;
  origSurvFsv: string;
  presBranchMv: string;
  presBranchFsv: string;
  presSurvMv: string;
  presSurvFsv: string;
  presAvMv: string;
  presAvFsv: string;
}

export interface Section9Row {
  sl: string;
  referenceNo: string;
  date: string;
  limit: string;
  remarks: string;
}

export interface Section10Particulars {
  nameOfClient: string;
  uniqueCustomerId: string;
  awcdMsdAccount: string;
  businessTypeNature: string;
  businessManagementType: string;
  clientType: string;
  keyPersonPhone: string;
  address: string;
  officeBusinessAddress: string;
  factoryAddress: string;
  purposeOfInvestment: string;
  limitOfInvestment: string;
  modeOfInvestment: string;
  lastVisitDate: string;
  stockPosition: string;
  fixedAsset: string;
  businessNetWorth: string;
  insuranceCoverage: string;
  tradeLicenseValidity: string;
  otherRegulatoryDocuments: string;
  dateOfOriginalInvestment: string;
  firstDisbursementManager: string;
  firstDateOfClassification: string;
}

export interface Section10DirectorRow {
  name: string;
  age: string;
  education: string;
  designation: string;
  sharePct: string;
  experience: string;
  netWorth: string;
}

export interface Section10BankDirectorRow {
  name: string;
  bankNbfi: string;
  designation: string;
  sharePct: string;
  numShares: string;
}

export interface Section11Row {
  year: string;
  grossSales: string;
  grossProfit: string;
  netProfit: string;
  totalAsset: string;
  totalLiabilities: string;
  netWorth: string;
}

export interface BranchProposalData {
  referenceNo: string; // #64#
  proposalDate: string; // #65#
  branchName: string;
  accountNo: string;
  subjectText: string;
  limitOriginal: string; // #6#
  limitRSD: string;
  securityMV: string; // #33#
  securityFSV: string; // #34#
  justificationText: string;
  previousBoardDecisions: string;
  natureOfBusiness: string; // #54#
  sec2Limit: string; // #56#
  sec2Validity: string; // #58#
  sec2Overdue: string;
  sec2Recovery: string;
  sec2Remarks: string;
  section3AsOnDate: string;
  section3Rows: Section3Row[];
  section4CibDate: string;
  section4Rows: Section4Row[];
  section5StatusNote: string;
  section5Rows: Section5Row[];
  section6StatusNote: string;
  section6Rows: Section6Row[];
  section7PrimaryNote: string;
  section7PrimaryRows: Section7PrimaryRow[];
  section7PrimaryTotal: string;
  section7CollateralNote: string;
  section7CollateralSchedules: string[];
  section7ValuationRows: Section7ValuationRow[];
  section7CoveragePct: string;
  section7Others: string;
  section8SourceOfRepayment: string;
  section9Rows: Section9Row[];
  section10Particulars: Section10Particulars;
  section10Directors: Section10DirectorRow[];
  section10BankDirectors: Section10BankDirectorRow[];
  section11FinancialsNote: string;
  section11Rows: Section11Row[];
  section12BranchVisitReport: string;
  section15BranchBicRec: string;
  section16BranchRec: string;
}

export interface OutstandingData {
  princOutBDT: number | null; // #7#
  chargeBDT: number | null; // #8#
  profitOutBDT: number | null; // #10#
  compensationBDT: number | null; // #11#
  unappliedProBDT: number | null; // #12#
  netOutBDT: number | null; // #25#
}

export interface PartyPerformanceData {
  productName: string; // #5# and #55#
  recoveryAmtBDT: number | null; // #60#
  printDate: string; // #68#
}

export interface OverdueData {
  overdueAmountBDT: number | null; // #59#
}

export interface AccountStatementData {
  accountNumber: string; // #2#
  acNo: string; // #36#
  branchName: string; // #3# (2nd line of doc)
  customerId: string; // #69#
  lastBalanceBDT: number | null; // #35#
}

export interface CibReportData {
  cibDate: string; // #62#
}

export interface CostOfFundData {
  surplusDeficitBDT: number | null; // #31#
}

export interface ManualInputsData {
  proposalDate: string; // #1#
  waiverAmount: string; // #4#
  settlementAmount: string; // #27#
  settlementPeriod: string; // #28#
  waiverRecClientBranch: string; // #45#
  waiverRecHO: string; // #47#
  sourceOfRepayment: string; // Section 8
}

export type ScrutinySectionKey =
  | 'justification'
  | 'sec1NatureOfBusiness'
  | 'sec2LiabilityOurBank'
  | 'sec3LiabilityOtherUnits'
  | 'sec4CibStatus'
  | 'sec5LegalAction'
  | 'sec6AuditCompliance'
  | 'sec7Security'
  | 'sec8SourceOfRepayment'
  | 'sec9HoSanctions'
  | 'sec10ClientParticulars'
  | 'sec11Financials'
  | 'clientPrayer'
  | 'branchVisit'
  | 'sec13BranchRecommendation'
  | 'sec14IccdRecommendation'
  | 'samdObservation';

export interface BranchProposalSource {
  sourceId?: string;
  fileName: string;
  rawText: string;
  pageImages: Array<{ mimeType: string; data: string }>;
}

export interface InsufficientSectionInfo {
  section: ScrutinySectionKey;
  reason: string;
}

export interface BatchScrutinyState {
  running: boolean;
  total: number;
  targetSections: ScrutinySectionKey[];
  completedSections: ScrutinySectionKey[];
  retryingSections: ScrutinySectionKey[];
  failedSections: Array<{ section: ScrutinySectionKey; error: string }>;
  insufficientSections: InsufficientSectionInfo[];
  skippedManualSections: ScrutinySectionKey[];
}

export interface Section2AiData {
  limit: string;
  validity: string;
  remarks: string;
  nonFundedNote: string;
  summary: string;
}

export interface Section3AiData {
  rows: Section3Row[];
  summary: string;
}

export interface Section4AiData {
  cibDate: string;
  rows: Section4Row[];
  summary: string;
}

export interface Section5AiData {
  statusNote: string;
  rows: Section5Row[];
  summary: string;
}

export interface Section6AiData {
  statusNote: string;
  rows: Section6Row[];
  summary: string;
}

export interface Section7AiData {
  primaryNote: string;
  primaryRows: Section7PrimaryRow[];
  primaryTotal: string;
  collateralNote: string;
  collateralSchedules: string[];
  valuationRows: Section7ValuationRow[];
  coveragePct: string;
  othersNote: string;
  summary: string;
}

export interface Section9AiData {
  rows: Section9Row[];
  summary: string;
}

export interface Section10AiData {
  particulars: Section10Particulars;
  directors: Section10DirectorRow[];
  bankDirectors: Section10BankDirectorRow[];
  summary: string;
}

export interface Section11AiData {
  financialsNote: string;
  rows: Section11Row[];
  summary: string;
}

export interface AiSectionsData {
  justification: string | null; // #51#
  sec1NatureOfBusiness: string | null; // Section 1 (#54#)
  sec2LiabilityOurBank: Section2AiData | null; // Section 2
  sec3LiabilityOtherUnits: Section3AiData | null; // Section 3
  sec4CibStatus: Section4AiData | null; // Section 4
  sec5LegalAction: Section5AiData | null; // Section 5
  sec6AuditCompliance: Section6AiData | null; // Section 6
  sec7Security: Section7AiData | null; // Section 7
  sec8SourceOfRepayment: string | null; // Section 8
  sec9HoSanctions: Section9AiData | null; // Section 9
  sec10ClientParticulars: Section10AiData | null; // Section 10
  sec11Financials: Section11AiData | null; // Section 11
  clientPrayer: string | null; // 12. Contents of the Client's Prayer
  branchVisit: string | null; // 12. Branch Visit Report
  sec13BranchRecommendation: string | null; // Section 13
  sec14IccdRecommendation: string | null; // Section 14
  samdObservation: string | null; // Section 15
}

export interface ExtractedDocumentsState {
  branchProposal: BranchProposalData | null;
  outstanding: OutstandingData | null;
  partyPerformance: PartyPerformanceData | null;
  overdue: OverdueData | null;
  accountStatement: AccountStatementData | null;
  cibReport: CibReportData | null;
  costOfFund: CostOfFundData | null;
}

export interface MissingFieldItem {
  code: string;
  label: string;
  source: string;
  reason: string;
}

export interface ResolvedICMemo {
  placeholders: Record<string, string>;
  missingFields: MissingFieldItem[];
  sourceOfRepayment: string;
  justificationLines: string; // #51#
  clientPrayer: string;
  branchVisitReport: string;
  samdObservation: string;
  sec13BranchRecommendation: string;
  sec14IccdRecommendation: string;
  sectionSummaries: Partial<Record<ScrutinySectionKey, string>>;
  sec2NonFundedNote: string;
  sec7OthersNote: string;
  branchProposal: BranchProposalData | null;
}

export type StageMemoDocType = 'mec' | 'board' | 'sanction';

export type StageScrutinySectionKey =
  | 'mecIcRecommendation'
  | 'boardMecRecommendation'
  | 'sanctionCrossVerifiedSummary';

export interface StageMemoInputs {
  icMemoNumber: string;
  icMemoDate: string;
  mecMemoNumber: string;
  mecMemoDate: string;
  boardMemoNumber: string;
  boardMemoDate: string;
  sanctionRefNumber: string;
  sanctionDate: string;
}

export type StageAiSectionsData = Record<StageScrutinySectionKey, string | null>;

export interface StageBatchScrutinyState {
  running: boolean;
  docType: StageMemoDocType | null;
  total: number;
  targetSections: StageScrutinySectionKey[];
  completedSections: StageScrutinySectionKey[];
  retryingSections: StageScrutinySectionKey[];
  failedSections: Array<{ section: StageScrutinySectionKey; error: string }>;
  insufficientSections: Array<{ section: StageScrutinySectionKey; reason: string }>;
  skippedManualSections: StageScrutinySectionKey[];
}

export interface ResolvedStageMemo {
  docType: StageMemoDocType;
  icMemo: ResolvedICMemo;
  stageInputs: StageMemoInputs;
  stageSections: Record<StageScrutinySectionKey, string>;
}

export interface MemoEngineDefinition {
  id: Exclude<MemoTabKey, 'manual'>;
  name: string;
  status: 'implemented' | 'coming_soon';
  description: string;
}

export type MemoDocType = Exclude<MemoTabKey, 'manual'>;

export interface TemplatePreviewBlock {
  type: 'paragraph' | 'table';
  text?: string;
  rows?: string[][];
}

export interface MemoTemplateRecord {
  docType: MemoDocType;
  fileName: string;
  fileSize: number;
  uploadedAt: string;
  bytes: Uint8Array;
  detectedTokens: string[];
  previewSections: TemplatePreviewBlock[];
}

export interface AppliedSyncState {
  appliedAt: string | null;
  version: number;
  memo: ResolvedICMemo;
  stageInputs: StageMemoInputs;
  stageAiSections: StageAiSectionsData;
}

export interface StageSourceAuditResult {
  summary: string;
  verifiedPoints: string[];
  discrepanciesOrMissing: string[];
}

export interface StageSourceUploadState {
  uploaded: boolean;
  fileName: string;
  sourceType: 'ic_memo' | 'mec_memo' | 'board_memo';
  processing: boolean;
  stage: string | null;
  error: string | null;
  extractedAt: string | null;
  extractedFieldCount: number;
  audit: StageSourceAuditResult | null;
}
