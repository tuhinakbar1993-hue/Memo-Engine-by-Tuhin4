import { createEmptyBranchProposal } from '../parsers/branchProposalParser';
import type {
  AiSectionsData,
  BranchProposalData,
  ExtractedDocumentsState,
  ManualInputsData,
  MissingFieldItem,
  ResolvedICMemo,
  ResolvedStageMemo,
  ScrutinySectionKey,
  StageAiSectionsData,
  StageMemoDocType,
  StageMemoInputs,
  StageScrutinySectionKey,
} from '../types/memo';
import { bdtToMillionStr, formatMillionNumber, parseIndianNumber } from '../utils/numberFormat';

export interface PlaceholderSpec {
  code: string;
  label: string;
  source: string;
}

export const IC_PLACEHOLDER_SPECS: PlaceholderSpec[] = [
  { code: '#1#', label: 'Proposal Date', source: 'Manual Inputs' },
  { code: '#2#', label: 'Account Number / Name of Concern', source: 'Account Statement (Account number)' },
  { code: '#3#', label: 'Branch Name', source: 'Account Statement (2nd line)' },
  { code: '#4#', label: 'Waiver Amount (Million BDT)', source: 'Manual Inputs' },
  { code: '#5#', label: 'Mode / Product Name', source: 'Party Performance (Product Name)' },
  { code: '#6#', label: 'Original Limit (Million BDT)', source: 'Branch Proposal (Main Table)' },
  { code: '#7#', label: 'Principal Outstanding (Million BDT)', source: 'Outstanding (Grand Total -> Princ Out)' },
  { code: '#8#', label: 'Charges (Million BDT)', source: 'Outstanding (Grand Total -> Charge)' },
  { code: '#10#', label: 'Charged Profit (Million BDT)', source: 'Outstanding (Grand Total -> Profit Out)' },
  { code: '#11#', label: 'Compensation (Million BDT)', source: 'Outstanding (Grand Total -> Compensation)' },
  { code: '#12#', label: 'Unapplied / Uncharged Comp. (Million BDT)', source: 'Outstanding (Grand Total -> Unapplied Pro)' },
  { code: '#25#', label: 'NET Outstanding (Million BDT)', source: 'Outstanding (Grand Total -> NET Out)' },
  { code: '#27#', label: 'Settlement Amount (Million BDT)', source: 'Manual Inputs' },
  { code: '#28#', label: 'Settlement Period / Deadline', source: 'Manual Inputs' },
  { code: '#31#', label: 'Cost of Fund Surplus/Deficit (Million BDT)', source: 'Cost of Fund (Excel)' },
  { code: '#32#', label: 'Breakup Total (#7# + #8# + #31#)', source: 'Calculated' },
  { code: '#33#', label: 'Security Value MV', source: 'Branch Proposal (Remarks -> MV)' },
  { code: '#34#', label: 'Security Value FSV', source: 'Branch Proposal (Remarks -> FSV)' },
  { code: '#35#', label: 'Down Payment / Last Balance (Million BDT)', source: 'Account Statement (Last Balance)' },
  { code: '#36#', label: 'Down Payment A/C No.', source: 'Account Statement (A/C No.)' },
  { code: '#42#', label: 'Total Profit+Comp+Unapplied (#10#+#8#+#11#+#12#)', source: 'Calculated' },
  { code: '#44#', label: 'Max Waiver Considering COF (#42# - #31#)', source: 'Calculated' },
  { code: '#45#', label: 'Waiver Recommended by Client & Branch', source: 'Manual Inputs' },
  { code: '#47#', label: 'Waiver Recommended by H/O', source: 'Manual Inputs' },
  { code: '#51#', label: 'Justification of Proposed Facility (3 lines)', source: 'AI Generate / Manual' },
  { code: '#54#', label: 'Nature of Business', source: 'Branch Proposal (Section 1)' },
  { code: '#55#', label: 'Sec. 2 Mode of Investment', source: 'Party Performance (Product Name)' },
  { code: '#56#', label: 'Sec. 2 Limit (Million BDT)', source: 'Branch Proposal (Section 2)' },
  { code: '#58#', label: 'Sec. 2 Validity', source: 'Branch Proposal (Section 2)' },
  { code: '#59#', label: 'Overdue Amount (Million BDT)', source: 'Overdue (Grand Total)' },
  { code: '#60#', label: 'Recovery Till Date (Million BDT)', source: 'Party Performance (Customer Total)' },
  { code: '#62#', label: 'CIB Report Date', source: 'CIB Report (Date of Inquiry)' },
  { code: '#64#', label: 'Branch Proposal Reference No.', source: 'Branch Proposal (Ref. No.)' },
  { code: '#65#', label: 'Branch Proposal Date', source: 'Branch Proposal (Proposal Date)' },
  { code: '#68#', label: 'O/S As On / Print Date', source: 'Party Performance (Print Date)' },
  { code: '#69#', label: 'Customer ID / CIS', source: 'Account Statement (Customer ID)' },
];

export function resolveICMemo(
  docs: ExtractedDocumentsState,
  manual: ManualInputsData,
  ai: AiSectionsData,
  overrides: Record<string, string>,
  manualEdited: Partial<Record<ScrutinySectionKey, boolean>> = {}
): ResolvedICMemo {
  const bp = docs.branchProposal;
  const out = docs.outstanding;
  const pp = docs.partyPerformance;
  const ovd = docs.overdue;
  const stmt = docs.accountStatement;
  const cib = docs.cibReport;
  const cof = docs.costOfFund;

  const hasStructuredAi = Boolean(
    ai.sec2LiabilityOurBank ||
      ai.sec3LiabilityOtherUnits ||
      ai.sec4CibStatus ||
      ai.sec5LegalAction ||
      ai.sec6AuditCompliance ||
      ai.sec7Security ||
      ai.sec9HoSanctions ||
      ai.sec10ClientParticulars ||
      ai.sec11Financials
  );

  let mergedBp: BranchProposalData | null = bp
    ? {
        ...bp,
        section10Particulars: { ...bp.section10Particulars },
      }
    : hasStructuredAi
      ? createEmptyBranchProposal()
      : null;

  if (mergedBp) {
    if (ai.sec2LiabilityOurBank) {
      if (manualEdited.sec2LiabilityOurBank || ai.sec2LiabilityOurBank.limit.trim()) {
        mergedBp.sec2Limit = ai.sec2LiabilityOurBank.limit.trim();
      }
      if (manualEdited.sec2LiabilityOurBank || ai.sec2LiabilityOurBank.validity.trim()) {
        mergedBp.sec2Validity = ai.sec2LiabilityOurBank.validity.trim();
      }
      if (manualEdited.sec2LiabilityOurBank || ai.sec2LiabilityOurBank.remarks.trim()) {
        mergedBp.sec2Remarks = ai.sec2LiabilityOurBank.remarks.trim();
      }
    }
    if (ai.sec3LiabilityOtherUnits) {
      mergedBp.section3Rows = ai.sec3LiabilityOtherUnits.rows;
    }
    if (ai.sec4CibStatus) {
      mergedBp.section4Rows = ai.sec4CibStatus.rows;
      if (manualEdited.sec4CibStatus || ai.sec4CibStatus.cibDate.trim()) {
        mergedBp.section4CibDate = ai.sec4CibStatus.cibDate.trim();
      }
    }
    if (ai.sec5LegalAction) {
      mergedBp.section5StatusNote = ai.sec5LegalAction.statusNote;
      mergedBp.section5Rows = ai.sec5LegalAction.rows;
    }
    if (ai.sec6AuditCompliance) {
      mergedBp.section6StatusNote = ai.sec6AuditCompliance.statusNote;
      mergedBp.section6Rows = ai.sec6AuditCompliance.rows;
    }
    if (ai.sec7Security) {
      mergedBp.section7PrimaryNote = ai.sec7Security.primaryNote;
      mergedBp.section7PrimaryRows = ai.sec7Security.primaryRows;
      mergedBp.section7PrimaryTotal = ai.sec7Security.primaryTotal;
      mergedBp.section7CollateralNote = ai.sec7Security.collateralNote;
      mergedBp.section7CollateralSchedules = ai.sec7Security.collateralSchedules;
      mergedBp.section7ValuationRows = ai.sec7Security.valuationRows;
      mergedBp.section7CoveragePct = ai.sec7Security.coveragePct;
      mergedBp.section7Others = ai.sec7Security.othersNote;
    }
    if (ai.sec9HoSanctions) {
      mergedBp.section9Rows = ai.sec9HoSanctions.rows;
    }
    if (ai.sec10ClientParticulars) {
      mergedBp.section10Particulars = { ...ai.sec10ClientParticulars.particulars };
      mergedBp.section10Directors = ai.sec10ClientParticulars.directors;
      mergedBp.section10BankDirectors = ai.sec10ClientParticulars.bankDirectors;
    }
    if (ai.sec11Financials) {
      mergedBp.section11FinancialsNote = ai.sec11Financials.financialsNote;
      mergedBp.section11Rows = ai.sec11Financials.rows;
    }
  }

  const base: Record<string, string> = {
    // 5.1 Manual Inputs
    '#1#': manual.proposalDate.trim(),
    '#4#': manual.waiverAmount.trim(),
    '#27#': manual.settlementAmount.trim(),
    '#28#': manual.settlementPeriod.trim(),
    '#45#': manual.waiverRecClientBranch.trim(),
    '#47#': manual.waiverRecHO.trim(),

    // 5.2 Account Statement
    '#2#': stmt?.accountNumber?.trim() || '',
    '#36#': stmt?.acNo?.trim() || '',
    '#3#': stmt?.branchName?.trim() || '',
    '#69#': stmt?.customerId?.trim() || '',
    '#35#': stmt ? bdtToMillionStr(stmt.lastBalanceBDT) : '',

    // 5.3 Outstanding (Grand Total, BDT -> Million BDT)
    '#7#': out ? bdtToMillionStr(out.princOutBDT) : '',
    '#8#': out ? bdtToMillionStr(out.chargeBDT) : '',
    '#10#': out ? bdtToMillionStr(out.profitOutBDT) : '',
    '#11#': out ? bdtToMillionStr(out.compensationBDT) : '',
    '#12#': out ? bdtToMillionStr(out.unappliedProBDT) : '',
    '#25#': out ? bdtToMillionStr(out.netOutBDT) : '',

    // 5.4 Party Performance
    '#5#': pp?.productName?.trim() || '',
    '#55#': pp?.productName?.trim() || '',
    '#60#': pp ? bdtToMillionStr(pp.recoveryAmtBDT) : '',
    '#68#': pp?.printDate?.trim() || '',

    // 5.5 Overdue
    '#59#': ovd ? bdtToMillionStr(ovd.overdueAmountBDT) : '',

    // 5.6 CIB Report (authoritative CIB Report first, fallback to Branch Proposal sec 4 or manual Sec 4 edit)
    '#62#':
      manualEdited.sec4CibStatus && ai.sec4CibStatus
        ? ai.sec4CibStatus.cibDate.trim() || cib?.cibDate?.trim() || ''
        : cib?.cibDate?.trim() || ai.sec4CibStatus?.cibDate?.trim() || bp?.section4CibDate?.trim() || '',

    // 5.7 Cost of Fund
    '#31#': cof ? bdtToMillionStr(cof.surplusDeficitBDT) : '',

    // 5.8 Branch Proposal (already in Million BDT — copy as written)
    '#6#': bp?.limitOriginal?.trim() || '',
    '#33#': bp?.securityMV?.trim() || '',
    '#34#': bp?.securityFSV?.trim() || '',
    '#54#':
      ai.sec1NatureOfBusiness !== null
        ? ai.sec1NatureOfBusiness.trim()
        : bp?.natureOfBusiness?.trim() || '',
    '#56#': mergedBp?.sec2Limit?.trim() || '',
    '#58#': mergedBp?.sec2Validity?.trim() || '',
    '#64#': bp?.referenceNo?.trim() || '',
    '#65#': bp?.proposalDate?.trim() || '',

    // 6. AI Generate (#51#)
    '#51#': ai.justification !== null ? ai.justification.trim() : '',
  };

  // Apply non-calculated manual overrides first so calculated fields react to user edits
  const resolved: Record<string, string> = { ...base };
  for (const [key, val] of Object.entries(overrides)) {
    if (key !== '#32#' && key !== '#42#' && key !== '#44#' && val !== undefined) {
      resolved[key] = val.trim();
    }
  }

  // 5.9 Calculated fields
  const n7 = parseIndianNumber(resolved['#7#']);
  const n8 = parseIndianNumber(resolved['#8#']);
  const n10 = parseIndianNumber(resolved['#10#']);
  const n11 = parseIndianNumber(resolved['#11#']);
  const n12 = parseIndianNumber(resolved['#12#']);
  const n31 = parseIndianNumber(resolved['#31#']);

  resolved['#32#'] =
    n7 !== null && n8 !== null && n31 !== null ? formatMillionNumber(n7 + n8 + n31) : '';

  resolved['#42#'] =
    n10 !== null && n8 !== null && n11 !== null && n12 !== null
      ? formatMillionNumber(n10 + n8 + n11 + n12)
      : '';

  const n42 = parseIndianNumber(resolved['#42#']);
  resolved['#44#'] =
    n42 !== null && n31 !== null ? formatMillionNumber(n42 - n31) : '';

  // Allow direct override on calculated fields if explicitly edited
  for (const calcKey of ['#32#', '#42#', '#44#']) {
    if (overrides[calcKey] !== undefined && overrides[calcKey].trim() !== '') {
      resolved[calcKey] = overrides[calcKey].trim();
    }
  }

  // Collect only unfilled placeholders for Validation Panel
  const missingFields: MissingFieldItem[] = [];
  for (const spec of IC_PLACEHOLDER_SPECS) {
    const val = resolved[spec.code];
    if (!val || val.trim() === '') {
      missingFields.push({
        code: spec.code,
        label: spec.label,
        source: spec.source,
        reason:
          spec.code === '#51#'
            ? 'Not generated yet (click "Scrutinize and Fill Up Through AI" in IC Memo or enter manually).'
            : `Missing or unreadable from ${spec.source}.`,
      });
    }
  }

  const sectionSummaries: Partial<Record<ScrutinySectionKey, string>> = {};
  if (ai.sec2LiabilityOurBank?.summary?.trim()) {
    sectionSummaries.sec2LiabilityOurBank = ai.sec2LiabilityOurBank.summary.trim();
  }
  if (ai.sec3LiabilityOtherUnits?.summary?.trim()) {
    sectionSummaries.sec3LiabilityOtherUnits = ai.sec3LiabilityOtherUnits.summary.trim();
  }
  if (ai.sec4CibStatus?.summary?.trim()) {
    sectionSummaries.sec4CibStatus = ai.sec4CibStatus.summary.trim();
  }
  if (ai.sec5LegalAction?.summary?.trim()) {
    sectionSummaries.sec5LegalAction = ai.sec5LegalAction.summary.trim();
  }
  if (ai.sec6AuditCompliance?.summary?.trim()) {
    sectionSummaries.sec6AuditCompliance = ai.sec6AuditCompliance.summary.trim();
  }
  if (ai.sec7Security?.summary?.trim()) {
    sectionSummaries.sec7Security = ai.sec7Security.summary.trim();
  }
  if (ai.sec9HoSanctions?.summary?.trim()) {
    sectionSummaries.sec9HoSanctions = ai.sec9HoSanctions.summary.trim();
  }
  if (ai.sec10ClientParticulars?.summary?.trim()) {
    sectionSummaries.sec10ClientParticulars = ai.sec10ClientParticulars.summary.trim();
  }
  if (ai.sec11Financials?.summary?.trim()) {
    sectionSummaries.sec11Financials = ai.sec11Financials.summary.trim();
  }

  return {
    placeholders: resolved,
    missingFields,
    sourceOfRepayment:
      manual.sourceOfRepayment.trim() || (ai.sec8SourceOfRepayment ?? '').trim(),
    justificationLines: resolved['#51#'] || '',
    clientPrayer: ai.clientPrayer !== null ? ai.clientPrayer.trim() : '',
    branchVisitReport: ai.branchVisit !== null ? ai.branchVisit.trim() : '',
    samdObservation: ai.samdObservation !== null ? ai.samdObservation.trim() : '',
    sec13BranchRecommendation:
      ai.sec13BranchRecommendation !== null ? ai.sec13BranchRecommendation.trim() : '',
    sec14IccdRecommendation:
      ai.sec14IccdRecommendation !== null ? ai.sec14IccdRecommendation.trim() : '',
    sectionSummaries,
    sec2NonFundedNote: ai.sec2LiabilityOurBank?.nonFundedNote?.trim() || '',
    sec7OthersNote: ai.sec7Security?.othersNote?.trim() || '',
    branchProposal: mergedBp,
  };
}

export const INITIAL_STAGE_INPUTS: StageMemoInputs = {
  icMemoNumber: '',
  icMemoDate: '',
  mecMemoNumber: '',
  mecMemoDate: '',
  boardMemoNumber: '',
  boardMemoDate: '',
  sanctionRefNumber: '',
  sanctionDate: '',
};

export const INITIAL_STAGE_AI_SECTIONS: StageAiSectionsData = {
  mecIcRecommendation: null,
  boardMecRecommendation: null,
  sanctionCrossVerifiedSummary: null,
};

export const STAGE_SECTIONS_BY_DOC: Record<StageMemoDocType, StageScrutinySectionKey[]> = {
  mec: ['mecIcRecommendation'],
  board: ['boardMecRecommendation'],
  sanction: ['sanctionCrossVerifiedSummary'],
};

export const STAGE_SECTION_LABELS: Record<StageScrutinySectionKey, string> = {
  mecIcRecommendation: 'Recommendation of the Investment Committee (IC)',
  boardMecRecommendation: 'Management Executive Committee (MEC) Recommendation',
  sanctionCrossVerifiedSummary: 'AI Cross-Verified Facility & Liability Confirmation',
};

export function formatStageOrder26(rawOrder: string, fallbackToken: string): string {
  const clean = rawOrder.trim();
  if (!clean) return `${fallbackToken}/26`;
  return /\/26$/i.test(clean) ? clean : `${clean}/26`;
}

export function buildMecSection16Paragraph(icMemoNumber: string, icMemoDate: string): string {
  const order = formatStageOrder26(icMemoNumber, '[Chronologic Order of the IC Memo]');
  const dt = icMemoDate.trim() || '[IC Memo Date]';
  return `This proposal was placed before the Investment Committee (IC) in its ${order} meeting held on ${dt}. The Committee recommended the proposal as presented.`;
}

export function buildBoardSection17Paragraph(mecMemoNumber: string, mecMemoDate: string): string {
  const order = formatStageOrder26(mecMemoNumber, '[Chronologic Order of the MEC Memo]');
  const dt = mecMemoDate.trim() || '[MEC Memo Date]';
  return `This proposal was placed before the Management Executive Committee (MEC) in its ${order} meeting held on ${dt}. The Committee recommended the proposal as presented.`;
}

export function resolveStageMemo(
  docType: StageMemoDocType,
  icMemo: ResolvedICMemo,
  stageInputs: StageMemoInputs,
  stageAi: StageAiSectionsData,
  stageManualEditedSections?: Partial<Record<StageScrutinySectionKey, boolean>>
): ResolvedStageMemo {
  const cleanInputs: StageMemoInputs = {
    icMemoNumber: stageInputs.icMemoNumber.trim(),
    icMemoDate: stageInputs.icMemoDate.trim(),
    mecMemoNumber: stageInputs.mecMemoNumber.trim(),
    mecMemoDate: stageInputs.mecMemoDate.trim(),
    boardMemoNumber: stageInputs.boardMemoNumber.trim(),
    boardMemoDate: stageInputs.boardMemoDate.trim(),
    sanctionRefNumber: stageInputs.sanctionRefNumber.trim(),
    sanctionDate: stageInputs.sanctionDate.trim(),
  };

  const manualMecIc = Boolean(stageManualEditedSections?.mecIcRecommendation);
  const manualBoardMec = Boolean(stageManualEditedSections?.boardMecRecommendation);

  const mecIcText =
    manualMecIc && (stageAi.mecIcRecommendation ?? '').trim()
      ? (stageAi.mecIcRecommendation ?? '').trim()
      : buildMecSection16Paragraph(cleanInputs.icMemoNumber, cleanInputs.icMemoDate);

  const boardMecText =
    manualBoardMec && (stageAi.boardMecRecommendation ?? '').trim()
      ? (stageAi.boardMecRecommendation ?? '').trim()
      : buildBoardSection17Paragraph(cleanInputs.mecMemoNumber, cleanInputs.mecMemoDate);

  const stageSections: Record<StageScrutinySectionKey, string> = {
    mecIcRecommendation: mecIcText,
    boardMecRecommendation: boardMecText,
    sanctionCrossVerifiedSummary: (stageAi.sanctionCrossVerifiedSummary ?? '').trim(),
  };

  return {
    docType,
    icMemo,
    stageInputs: cleanInputs,
    stageSections,
  };
}

