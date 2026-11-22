import dotenv from 'dotenv';
import { createHash } from 'crypto';
import express from 'express';
import http from 'http';
import path from 'path';
import { fileURLToPath } from 'url';
import { GoogleGenAI, ThinkingLevel, Type, type Schema } from '@google/genai';

dotenv.config();
dotenv.config({ path: '.env.local', override: false });

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = 3000;

const MISSING_API_KEY_MESSAGE =
  'GEMINI_API_KEY is not configured. Please add or select a Gemini API key in Settings > Secrets and retry.';

class MissingApiKeyError extends Error {
  constructor(message = MISSING_API_KEY_MESSAGE) {
    super(message);
    this.name = 'MissingApiKeyError';
  }
}

let cachedAiClient: { apiKey: string; client: GoogleGenAI } | null = null;

function getAiClient(): GoogleGenAI {
  const apiKey = (process.env.GEMINI_API_KEY || process.env.API_KEY || '').trim();
  if (!apiKey) {
    throw new MissingApiKeyError();
  }
  if (cachedAiClient && cachedAiClient.apiKey === apiKey) {
    return cachedAiClient.client;
  }
  const client = new GoogleGenAI({
    apiKey,
    httpOptions: {
      headers: {
        'User-Agent': 'aistudio-build',
      },
    },
  });
  cachedAiClient = { apiKey, client };
  return client;
}

function resolveAiClient(): { ok: true; ai: GoogleGenAI } | { ok: false; status: number; error: string } {
  try {
    return { ok: true, ai: getAiClient() };
  } catch (error) {
    if (error instanceof MissingApiKeyError) {
      return { ok: false, status: 503, error: error.message };
    }
    return {
      ok: false,
      status: 500,
      error: error instanceof Error ? error.message : 'Failed to initialize Gemini AI client.',
    };
  }
}

type ScrutinySectionKey =
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

type StageMemoDocType = 'mec' | 'board' | 'sanction';

type StageScrutinySectionKey =
  | 'mecIcRecommendation'
  | 'boardMecRecommendation'
  | 'sanctionCrossVerifiedSummary';

function strField(description: string): Schema {
  return { type: Type.STRING, description };
}

function getScrutinyConfig(section: ScrutinySectionKey): {
  prompt: string;
  schema: Schema;
} {
  switch (section) {
    case 'justification':
      return {
        prompt: `Scrutinize the uploaded Branch Proposal (especially "Justification of The Proposed Facility", subject/terms, business/distress background, recovery/down payment details, and branch recommendations) together with verified memo figures below.
Write clean, formal banking text strictly for IC Memo item "Justification of The Proposed Facility:" ("a.").
Format as exactly 3 numbered justification points separated by newlines ("1. ...", "2. ...", "3. ..."), covering:
1. Client's current business/financial position and genuine distress reflected in the Branch Proposal.
2. Recovery and settlement rationale (down payment, principal/Cost of Fund recovery, and protection of bank interest) grounded strictly in the Branch Proposal and provided figures.
3. Immediate reduction/regularization of classified NPL exposure and avoidance of prolonged litigation.
CRITICAL CLEAN-TEXT RULES: Output ONLY formal banking justification points in "text". Do NOT include conversational openings, AI notes, or explanations ("Based on the uploaded document", "Here is...", etc.), and do NOT use markdown asterisks.
If the Branch Proposal lacks information needed for this section, include ONLY what is factually supported (or "" if none) in "text" and explain what is unavailable strictly inside "insufficientReason".`,
        schema: {
          type: Type.OBJECT,
          properties: {
            text: strField('Clean official banking justification points (1., 2., 3.) with no AI notes or disclaimers, or "" if source info is completely absent.'),
            insufficientReason: strField('Empty string "" if sufficiently supported; otherwise brief software UI note stating what info was unavailable in Branch Proposal.'),
          },
          required: ['text', 'insufficientReason'],
        },
      };

    case 'sec1NatureOfBusiness':
      return {
        prompt: `Scrutinize the uploaded Branch Proposal (Section "1. Nature of Business", Section 10 "Business Type/Nature", and business description) to populate IC Memo Section "1. Nature of Business:".
Return ONLY the concise official Nature of Business phrase in "text" (e.g., "Agriculture (Fisheries/Dairy)" or exact activity wording from the Branch Proposal).
Do NOT include conversational remarks, "Based on...", AI notes, or markdown asterisks in "text".
If the Branch Proposal does not state the nature of business, set "text" to "" and explain strictly in "insufficientReason".`,
        schema: {
          type: Type.OBJECT,
          properties: {
            text: strField('Exact concise Nature of Business label/phrase for the memo, or "" if unavailable.'),
            insufficientReason: strField('Empty string "" if found; otherwise concise UI notice explaining missing information.'),
          },
          required: ['text', 'insufficientReason'],
        },
      };

    case 'sec2LiabilityOurBank':
      return {
        prompt: `Scrutinize Section "2. Liability Position of the Clients and Its Sister Concerns with Our Bank" in the uploaded Branch Proposal.
Extract strictly what appears in the Branch Proposal:
- limit: Sanctioned limit in Million BDT (e.g. "0.50" or "0.50 Million"; "" if absent).
- validity: Validity / expiry date and status (e.g. "01.09.2025 (Exp.)"; "" if absent).
- remarks: Exact remarks in Section 2 table ("" if blank).
- nonFundedNote: Exact Non-Funded or sister concern note in Section 2 table ("" if blank/Nil).
- summary: For software UI audit log only — if key fields are blank/insufficient in the Branch Proposal, state concisely why; otherwise "". Never fabricate figures.`,
        schema: {
          type: Type.OBJECT,
          properties: {
            limit: strField('Section 2 Limit from Branch Proposal, or "" if absent.'),
            validity: strField('Section 2 Validity/Expiry from Branch Proposal, or "" if absent.'),
            remarks: strField('Section 2 Remarks from Branch Proposal, or "" if blank.'),
            nonFundedNote: strField('Non-Funded or sister concern remarks from Section 2, or "" if none.'),
            summary: strField('Software UI insufficient-data note only if info is missing/blank in Branch Proposal; else "".'),
          },
          required: ['limit', 'validity', 'remarks', 'nonFundedNote', 'summary'],
        },
      };

    case 'sec3LiabilityOtherUnits':
      return {
        prompt: `Scrutinize Section "3. Liability of the Client & Its Sister/Allied Concerns with Other Units of SIB" (or other banks/NBFIs) in the uploaded Branch Proposal.
Extract every row present in Section 3 of the Branch Proposal into "rows" (with exact columns: sl, nameOfConcern, bankOrNbfi, natureOfFacility, limit, outstanding, overdueEol, expiryDate, clStatus).
- If the Branch Proposal table has blank/empty rows or states "Nil" / "N/A", reflect that accurately (do NOT fabricate bank names, limits, or facilities; keep empty cells as "" or "Nil" only where the Branch Proposal actually states Nil).
- In "summary", write a brief formal statement explaining what the Branch Proposal shows for Section 3 (e.g., whether any sister/allied or other-unit/other-bank liabilities exist, or explicitly stating that the Branch Proposal records Nil / leaves the table blank / provides insufficient information).`,
        schema: {
          type: Type.OBJECT,
          properties: {
            rows: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  sl: strField('SL (e.g. a., b., c., Sub-Total, Grand Total)'),
                  nameOfConcern: strField('Name of Concern'),
                  bankOrNbfi: strField('Name of the Banks/NBFIs'),
                  natureOfFacility: strField('Nature of Facility'),
                  limit: strField('Limit'),
                  outstanding: strField('Outstanding with profit/Interest'),
                  overdueEol: strField('Overdue/EOL'),
                  expiryDate: strField('Expiry Date'),
                  clStatus: strField('CL Status'),
                },
                required: [
                  'sl',
                  'nameOfConcern',
                  'bankOrNbfi',
                  'natureOfFacility',
                  'limit',
                  'outstanding',
                  'overdueEol',
                  'expiryDate',
                  'clStatus',
                ],
              },
            },
            summary: strField('Formal banking summary of Section 3 findings or note on Nil/insufficient info.'),
          },
          required: ['rows', 'summary'],
        },
      };

    case 'sec4CibStatus':
      return {
        prompt: `Scrutinize Section "4. CIB Status" in the uploaded Branch Proposal thoroughly.
Extract:
- cibDate: The CIB report date mentioned in Section 4 of the Branch Proposal (e.g. "17.11.2026"), or "" if absent.
- rows: Extract the exact values from the Section 4 CIB Status table for each status row (STD, SMA, SS, DF, BL, Willful Default (WD), Write-Off, Non-Funded, Stay Order, TOTAL) across columns: status, unitBorrower, unitGuarantor, otherBorrower, otherGuarantor, remarks.
- Fix obvious scanner/OCR typos ONLY when clearly verifiable from the visual table image (e.g. "Nit" -> "Nil", "10/-" if actually "0/-" or "10/-" per image — read the visual page image carefully!). Leave genuinely blank cells as "". Do NOT invent figures.
- summary: Write 1-2 formal banking sentences summarizing the CIB status as recorded in the Branch Proposal (or clearly state if CIB information is unavailable/insufficient in the Branch Proposal).`,
        schema: {
          type: Type.OBJECT,
          properties: {
            cibDate: strField('CIB report date from Branch Proposal Section 4, or empty string.'),
            rows: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  status: strField('Status label (STD, SMA, SS, DF, BL, Willful Default (WD), Write-Off, Non-Funded, Stay Order, TOTAL)'),
                  unitBorrower: strField('With respective unit - Borrower'),
                  unitGuarantor: strField('With respective unit - Guarantor'),
                  otherBorrower: strField('With Other Bank - Borrower'),
                  otherGuarantor: strField('With Other Bank - Guarantor'),
                  remarks: strField('Remarks'),
                },
                required: ['status', 'unitBorrower', 'unitGuarantor', 'otherBorrower', 'otherGuarantor', 'remarks'],
              },
            },
            summary: strField('Concise formal summary of CIB status from the Branch Proposal.'),
          },
          required: ['cibDate', 'rows', 'summary'],
        },
      };

    case 'sec5LegalAction':
      return {
        prompt: `Scrutinize Section "5. Legal Action Status as on Memo Date" in the uploaded Branch Proposal.
Extract:
- statusNote: Any header status note such as "Nil" or "N/A" next to the heading in the Branch Proposal (or "" if none).
- rows: Any legal case rows listed in Section 5 of the Branch Proposal (sl, typeOfCase, caseNo, filingDate, amount, previousDate, nextDate, presentStatus). If no cases are filed / table is empty, return an empty array []. Do NOT invent case numbers or dates.
- summary: A brief formal banking statement describing the legal action status strictly per the Branch Proposal (or clearly stating if legal action details are marked Nil/N/A or insufficient in the Branch Proposal).`,
        schema: {
          type: Type.OBJECT,
          properties: {
            statusNote: strField('Status note next to Section 5 heading (e.g. Nil, N/A, or empty string).'),
            rows: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  sl: strField('Sl'),
                  typeOfCase: strField('Type of Case'),
                  caseNo: strField('Case No.'),
                  filingDate: strField('Filing date'),
                  amount: strField('Amount'),
                  previousDate: strField('Previous date'),
                  nextDate: strField('Next date'),
                  presentStatus: strField('Present status'),
                },
                required: ['sl', 'typeOfCase', 'caseNo', 'filingDate', 'amount', 'previousDate', 'nextDate', 'presentStatus'],
              },
            },
            summary: strField('Formal banking summary of Legal Action Status from the Branch Proposal.'),
          },
          required: ['statusNote', 'rows', 'summary'],
        },
      };

    case 'sec6AuditCompliance':
      return {
        prompt: `Scrutinize Section "6. Audit, Compliance and Regulatory Authority Related Status" in the uploaded Branch Proposal.
Extract:
- statusNote: Any note next to the Section 6 heading (e.g., "Nil", "N/A", or "").
- rows: Extract the exact inspection dates, objection status (Yes/No), objections, and branch reply/current position for each authority row in Section 6 (a. IC & CD, b. DBI, Bangladesh Bank, c. DOS, Bangladesh Bank, d. BFIU, Bangladesh Bank, e. BRD, Bangladesh Bank, f. ACC(Anti-Corruption Commission)).
- Preserve exact values visible in the Branch Proposal image/text. Do NOT fabricate inspection dates or replies.
- summary: A concise formal banking summary of the audit, compliance, and regulatory inspection position in the Branch Proposal (or clearly note if insufficient/unavailable).`,
        schema: {
          type: Type.OBJECT,
          properties: {
            statusNote: strField('Status note next to Section 6 heading (e.g., N/A, Nil, or empty string).'),
            rows: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  sl: strField('SL'),
                  authority: strField('Name of the Authority'),
                  lastInspectionDate: strField('Last Inspection / Report Date'),
                  objectionStatus: strField('Objection Status (Yes/No)'),
                  objections: strField('Objections (If yes)'),
                  branchReply: strField('Branch Reply/ Current Position'),
                },
                required: ['sl', 'authority', 'lastInspectionDate', 'objectionStatus', 'objections', 'branchReply'],
              },
            },
            summary: strField('Formal summary of Audit, Compliance & Regulatory Status from Branch Proposal.'),
          },
          required: ['statusNote', 'rows', 'summary'],
        },
      };

    case 'sec7Security':
      return {
        prompt: `Scrutinize Section "7. Security" (including a. Primary, b. Collateral, schedule descriptions, valuation table, coverage %, and c. Others, plus MV/FSV in Remarks if noted) in the uploaded Branch Proposal.
Extract strictly what is recorded in the Branch Proposal:
- primaryNote: Note next to "a. Primary:" (e.g. "Nil", "N/A", or "").
- primaryRows: Rows in the Primary Security table (particulars, value).
- primaryTotal: Total value of primary security (or "" if blank/Nil).
- collateralNote: Note next to "b. Collateral:" (e.g. "N/A", "Nil", or "").
- collateralSchedules: List of collateral schedule descriptions if filled in the Branch Proposal (if template placeholders are blank/unfilled, return standard schedule lines only as shown or empty array if marked N/A).
- valuationRows: Rows in the Collateral Valuation table (schedule, origBranchMv, origBranchFsv, origSurvMv, origSurvFsv, presBranchMv, presBranchFsv, presSurvMv, presSurvFsv, presAvMv, presAvFsv). Leave unfilled cells as "".
- coveragePct: "Outstanding to Lowest Security value Coverage (%)" value if stated in the Branch Proposal (or "" if blank).
- othersNote: Any additional text under "c. Others:" if stated beyond standard charge documents.
- summary: A concise formal banking assessment of primary & collateral security strictly grounded in the Branch Proposal (explicitly stating if primary/collateral security is marked Nil/N/A or lacks valuation figures in the Branch Proposal).`,
        schema: {
          type: Type.OBJECT,
          properties: {
            primaryNote: strField('Note next to a. Primary (e.g. Nil, N/A, or empty).'),
            primaryRows: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  particulars: strField('Particulars'),
                  value: strField('Value'),
                },
                required: ['particulars', 'value'],
              },
            },
            primaryTotal: strField('Total value of primary security'),
            collateralNote: strField('Note next to b. Collateral (e.g. N/A, Nil, or empty).'),
            collateralSchedules: { type: Type.ARRAY, items: { type: Type.STRING } },
            valuationRows: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  schedule: strField('Schedule'),
                  origBranchMv: strField('Original Branch MV'),
                  origBranchFsv: strField('Original Branch FSV'),
                  origSurvMv: strField('Original Surveyor MV'),
                  origSurvFsv: strField('Original Surveyor FSV'),
                  presBranchMv: strField('Present Branch MV'),
                  presBranchFsv: strField('Present Branch FSV'),
                  presSurvMv: strField('Present Surveyor MV'),
                  presSurvFsv: strField('Present Surveyor FSV'),
                  presAvMv: strField('Present AV Cell MV'),
                  presAvFsv: strField('Present AV Cell FSV'),
                },
                required: [
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
                ],
              },
            },
            coveragePct: strField('Coverage percentage if stated, else empty string.'),
            othersNote: strField('Additional note under c. Others if any, else empty string.'),
            summary: strField('Concise formal banking summary of Section 7 Security in the Branch Proposal.'),
          },
          required: [
            'primaryNote',
            'primaryRows',
            'primaryTotal',
            'collateralNote',
            'collateralSchedules',
            'valuationRows',
            'coveragePct',
            'othersNote',
            'summary',
          ],
        },
      };

    case 'sec8SourceOfRepayment':
      return {
        prompt: `Scrutinize Section "8. Source of Repayment" (and repayment/settlement funding details across the Branch Proposal).
Write clean, official banking prose strictly stating the source of repayment for IC Memo Section "8. Source of Repayment:" based solely on the Branch Proposal.
Do NOT include conversational openings, "Based on the uploaded document", AI notes, or markdown asterisks in "text".
If the Branch Proposal contains insufficient or no information on the source of repayment, set "text" to "" and state what is missing in "insufficientReason".`,
        schema: {
          type: Type.OBJECT,
          properties: {
            text: strField('Clean formal Source of Repayment prose for the memo body, or "" if unavailable.'),
            insufficientReason: strField('Empty string "" if supported; otherwise brief UI note explaining missing information in Branch Proposal.'),
          },
          required: ['text', 'insufficientReason'],
        },
      };

    case 'sec9HoSanctions':
      return {
        prompt: `Scrutinize Section "9. Sequence of Head Office Sanction(S)" in the uploaded Branch Proposal.
Extract every sanction row visible in Section 9 (columns: sl, referenceNo, date, limit, remarks). Read the visual page image carefully so reference numbers, dates, and limits are accurate.
If rows are blank or no Head Office sanctions are listed, return only populated rows (or empty array if none) and clearly indicate in "summary" what the Branch Proposal shows (or that Head Office sanction details are unavailable/insufficient in the Branch Proposal). Do NOT invent reference numbers or dates.`,
        schema: {
          type: Type.OBJECT,
          properties: {
            rows: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  sl: strField('SL (a., b., c., etc.)'),
                  referenceNo: strField('HO Sanction Reference No.'),
                  date: strField('Date'),
                  limit: strField('Limit'),
                  remarks: strField('Remarks'),
                },
                required: ['sl', 'referenceNo', 'date', 'limit', 'remarks'],
              },
            },
            summary: strField('Concise formal summary of Head Office Sanction sequence or unavailable/insufficient note.'),
          },
          required: ['rows', 'summary'],
        },
      };

    case 'sec10ClientParticulars':
      return {
        prompt: `Scrutinize Section "10. Brief Particulars of The Client" (items a. through u., plus "Particular of the Proprietor/Partner/Directors" and "Remarks: Director of any Bank/ FI/ Insurance company") in the uploaded Branch Proposal.
Carefully read the visual page images and text to extract EVERY available item accurately without OCR errors:
- particulars: nameOfClient (a), uniqueCustomerId (b), awcdMsdAccount (c), businessTypeNature (d), businessManagementType (e), clientType (f), keyPersonPhone (g), address (g), officeBusinessAddress (h), factoryAddress (h), purposeOfInvestment (i), limitOfInvestment (j), modeOfInvestment (k), lastVisitDate (l), stockPosition (m), fixedAsset (n), businessNetWorth (o), insuranceCoverage (p), tradeLicenseValidity (q), otherRegulatoryDocuments (r), dateOfOriginalInvestment (s), firstDisbursementManager (t), firstDateOfClassification (u). Leave genuinely missing/blank items as "".
- directors: rows from "Particular of the Proprietor/Partner/Directors" table (empty [] if blank/N/A).
- bankDirectors: rows from "Remarks: Director of any Bank/ FI/ Insurance company" table (empty [] if blank/N/A).
- summary: Concise formal note summarizing key client particulars and explicitly noting if director/bank-director tables or specific particulars are marked N/A or unavailable in the Branch Proposal.`,
        schema: {
          type: Type.OBJECT,
          properties: {
            particulars: {
              type: Type.OBJECT,
              properties: {
                nameOfClient: strField('a. Name of the Client'),
                uniqueCustomerId: strField('b. Unique Customer ID'),
                awcdMsdAccount: strField('c. AWCD/MSD Account with opening date'),
                businessTypeNature: strField('d. Business Type/Nature'),
                businessManagementType: strField('e. Business Management Type'),
                clientType: strField('f. Client Type'),
                keyPersonPhone: strField('g. Key Person with Phone No.'),
                address: strField('g. Address'),
                officeBusinessAddress: strField('h. Office/Business Address'),
                factoryAddress: strField('h. Factory Address'),
                purposeOfInvestment: strField('i. Purpose of Investment'),
                limitOfInvestment: strField('j. Limit of Investment'),
                modeOfInvestment: strField('k. Mode of Investment'),
                lastVisitDate: strField('l. Last Visit Date'),
                stockPosition: strField('m. Stock Position'),
                fixedAsset: strField('n. Fixed Asset'),
                businessNetWorth: strField('o. Business Net Worth'),
                insuranceCoverage: strField('p. Insurance Coverage'),
                tradeLicenseValidity: strField('q. Trade License & Validity'),
                otherRegulatoryDocuments: strField('r. Other Regulatory Documents'),
                dateOfOriginalInvestment: strField('s. Date of Original Investment'),
                firstDisbursementManager: strField('t. 1st Disbursement Manager'),
                firstDateOfClassification: strField('u. First Date of Classification with Status'),
              },
              required: [
                'nameOfClient',
                'uniqueCustomerId',
                'awcdMsdAccount',
                'businessTypeNature',
                'businessManagementType',
                'clientType',
                'keyPersonPhone',
                'address',
                'officeBusinessAddress',
                'factoryAddress',
                'purposeOfInvestment',
                'limitOfInvestment',
                'modeOfInvestment',
                'lastVisitDate',
                'stockPosition',
                'fixedAsset',
                'businessNetWorth',
                'insuranceCoverage',
                'tradeLicenseValidity',
                'otherRegulatoryDocuments',
                'dateOfOriginalInvestment',
                'firstDisbursementManager',
                'firstDateOfClassification',
              ],
            },
            directors: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  name: strField('Name of the Directors'),
                  age: strField('Age'),
                  education: strField('Education'),
                  designation: strField('Designation'),
                  sharePct: strField('Share (%)'),
                  experience: strField('Experience'),
                  netWorth: strField('Net worth'),
                },
                required: ['name', 'age', 'education', 'designation', 'sharePct', 'experience', 'netWorth'],
              },
            },
            bankDirectors: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  name: strField('Name of the Director'),
                  bankNbfi: strField('Name of the bank/NBFIs'),
                  designation: strField('Designation/Position'),
                  sharePct: strField('Percentage of Shares'),
                  numShares: strField('Number of Shares'),
                },
                required: ['name', 'bankNbfi', 'designation', 'sharePct', 'numShares'],
              },
            },
            summary: strField('Concise formal summary of client particulars and any unavailable/N/A items.'),
          },
          required: ['particulars', 'directors', 'bankDirectors', 'summary'],
        },
      };

    case 'sec11Financials':
      return {
        prompt: `Scrutinize Section "11. Financials of The Client" in the uploaded Branch Proposal.
Extract:
- financialsNote: Any note such as "Nil" or "N/A" if the client has no business or financials are marked Nil in the Branch Proposal (or "" if actual figures exist).
- rows: Financial rows (year, grossSales, grossProfit, netProfit, totalAsset, totalLiabilities, netWorth) if figures appear in the Branch Proposal (empty [] if marked Nil/N/A/blank). Do NOT fabricate financial figures.
- summary: A brief formal banking statement summarizing the financial position reported in the Branch Proposal (or explicitly noting that financials are marked Nil / unavailable / insufficient in the Branch Proposal).`,
        schema: {
          type: Type.OBJECT,
          properties: {
            financialsNote: strField('Note such as Nil or N/A if stated, else empty string.'),
            rows: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  year: strField('Year'),
                  grossSales: strField('Gross Sales'),
                  grossProfit: strField('Gross Profit'),
                  netProfit: strField('Net Profit'),
                  totalAsset: strField('Total Asset'),
                  totalLiabilities: strField('Total liabilities'),
                  netWorth: strField('Net Worth'),
                },
                required: ['year', 'grossSales', 'grossProfit', 'netProfit', 'totalAsset', 'totalLiabilities', 'netWorth'],
              },
            },
            summary: strField('Formal banking summary of Client Financials or Nil/insufficient notice.'),
          },
          required: ['financialsNote', 'rows', 'summary'],
        },
      };

    case 'clientPrayer':
      return {
        prompt: `Scrutinize the uploaded Branch Proposal thoroughly for any details of the customer's written application / prayer ("Contents of the Client's Prayer", client's request letter/application, subject clause, requested waiver amount, offered settlement amount & down payment, proposed repayment period/timeline, stated business distress/reasons for default, and commitment for full and final exit).
Write clean, concise, official banking prose for IC Memo Section "12. Contents of the Client's Prayer:" strictly grounded in the Branch Proposal and verified memo figures below.
CRITICAL CLEAN-TEXT RULES: Output ONLY formal third-person banking memorandum prose describing the client's prayer/request in "text" (1-2 concise formal paragraphs). Do NOT include "Based on the uploaded document", conversational framing, AI notes, or markdown asterisks.
If the Branch Proposal contains insufficient information about the client's application/prayer, keep ONLY verified facts (or "" if none) in "text" and explain what is missing strictly inside "insufficientReason".`,
        schema: {
          type: Type.OBJECT,
          properties: {
            text: strField("Clean official banking prose summarizing Contents of the Client's Prayer, or '' if unavailable."),
            insufficientReason: strField('Empty string "" if supported; otherwise brief UI note listing missing client prayer details.'),
          },
          required: ['text', 'insufficientReason'],
        },
      };

    case 'branchVisit':
      return {
        prompt: `Scrutinize all relevant sections of the uploaded Branch Proposal (Section 12 "Branch Visit Report", Section 10 Client Particulars, office/factory addresses, last visit date, stock position, fixed assets, business net worth, Section 7 security details, and branch recommendations) to write Section "12. Branch Visit Report:" for the Head Office IC Memo.
CRITICAL REQUIREMENT: Write strictly in the THIRD PERSON regarding the branch (e.g., "The officials of [Branch Name] Branch...", "Branch management conducted a physical inspection...", never use "we", "us", "our", or "our branch").
Write 1 to 2 clean, formal banking paragraphs containing ONLY verified officials' names/designations, visit date, client/concern name, premises inspected, physical business/stock findings, and inspection outcome from the Branch Proposal.
Do NOT include "Based on the uploaded Branch Proposal", AI notes, disclaimers, or markdown asterisks inside "text".
If visit details are absent/insufficient in the Branch Proposal, keep ONLY verified facts (or "" if none) in "text" and record what is missing strictly in "insufficientReason".`,
        schema: {
          type: Type.OBJECT,
          properties: {
            text: strField('Clean third-person formal Branch Visit Report prose with no AI notes or disclaimers, or "" if unavailable.'),
            insufficientReason: strField('Empty string "" if sufficient visit facts were found; otherwise brief UI notice of missing visit details.'),
          },
          required: ['text', 'insufficientReason'],
        },
      };

    case 'sec13BranchRecommendation':
      return {
        prompt: `Scrutinize the uploaded Branch Proposal for the Branch's specific recommendation (Section 15 "Recommendation of Branch Investment Committee", Section 16 "Branch Recommendation", Proposal Reference No., Proposal Date, settlement amount, waiver amount, settlement period, and conditions).
Write clean formal banking prose for IC Memo Section "13. Recommendation of The Branch:" stating the exact Branch / BIC recommendation.
Do NOT include "Based on the uploaded document", conversational framing, AI commentary, or markdown asterisks inside "text".
If recommendation details are blank or insufficient in the Branch Proposal, set "text" to verified portions only (or "" if none) and record the gap strictly inside "insufficientReason".`,
        schema: {
          type: Type.OBJECT,
          properties: {
            text: strField('Clean formal Branch Recommendation prose for the memo body, or "" if absent.'),
            insufficientReason: strField('Empty string "" if sufficient; otherwise brief UI notice of missing recommendation details.'),
          },
          required: ['text', 'insufficientReason'],
        },
      };

    case 'sec14IccdRecommendation':
      return {
        prompt: `Scrutinize the uploaded Branch Proposal for any recommendation or comment from the Internal Control & Compliance Division (ICCD) (Section 14 or ICCD remarks).
Write clean formal content for IC Memo Section "14. Recommendation of Internal Control & Compliance Division:":
- If the Branch Proposal contains an ICCD recommendation, summarize it directly in clean formal banking prose without meta-commentary.
- If the Branch Proposal states "N/A", "Nil", or has no ICCD recommendation, return strictly "N/A" in "text" (no conversational notes or parenthetical AI explanations!).`,
        schema: {
          type: Type.OBJECT,
          properties: {
            text: strField('Clean ICCD Recommendation text or exact "N/A".'),
            insufficientReason: strField('Empty string "" unless Branch Proposal appears unreadable/missing.'),
          },
          required: ['text', 'insufficientReason'],
        },
      };

    case 'samdObservation':
      return {
        prompt: `Thoroughly scrutinize the uploaded Branch Proposal together with the verified IC Memo figures below, and write Section "15. Observation and Recommendation of SAMD:" as Head Office SAMD.
Write exactly five (5) substantive, formal banking points separated by newlines ("1. ...", "2. ...", "3. ...", "4. ...", "5. ...") addressing:
1. Verification of total outstanding liability breakup, classification/overdue/CIB status, and confirmation of absence of fraud, forgery, or willful default per records.
2. Assessment of primary & collateral security position/coverage, legal action status, and audit/regulatory compliance position.
3. Cost of Fund & waiver adequacy analysis (down payment, recovery vs. principal/charges/Cost of Fund, maximum permissible waiver considering COF, and conformity with BRPD Circular Letter No. 23 dated 29.06.2026).
4. Evaluation of the client's business/financial condition, Branch Visit findings, Branch Recommendation, and source of repayment.
5. Formal SAMD recommendation placing the Exit with Waiver proposal before the Investment Committee for exit within the stipulated settlement period subject to compliance with special terms and conditions.
CRITICAL CLEAN-TEXT RULES: Output ONLY crisp official SAMD memorandum points in "text". Do NOT add conversational remarks ("Here is...", "Based on the uploaded document..."), disclaimers, meta-commentary, or markdown asterisks. Ground every figure strictly in the Branch Proposal and context; record any missing source inputs in "insufficientReason".`,
        schema: {
          type: Type.OBJECT,
          properties: {
            text: strField('Clean official 5 numbered SAMD observation & recommendation points separated by newlines.'),
            insufficientReason: strField('Empty string "" if supported; otherwise brief UI note listing missing source inputs.'),
          },
          required: ['text', 'insufficientReason'],
        },
      };
  }
}

const VALID_SCRUTINY_SECTIONS = new Set<ScrutinySectionKey>([
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
]);

interface IncomingBranchProposalSource {
  sourceId?: string;
  fileName?: string;
  rawText?: string;
  pageImages?: Array<{ mimeType: string; data: string }>;
}

interface CachedBranchProposalSource {
  sourceId: string;
  fileName: string;
  rawText: string;
  imageParts: Array<{ inlineData: { mimeType: string; data: string } }>;
}

const MAX_CACHED_SOURCES = 20;
const branchProposalCache = new Map<string, CachedBranchProposalSource>();

function buildCachedSource(
  sourceId: string,
  incoming: IncomingBranchProposalSource
): CachedBranchProposalSource | null {
  const rawText = typeof incoming.rawText === 'string' ? incoming.rawText.slice(0, 75000) : '';
  const pageImages = Array.isArray(incoming.pageImages) ? incoming.pageImages.slice(0, 14) : [];
  const imageParts: Array<{ inlineData: { mimeType: string; data: string } }> = [];
  for (const img of pageImages) {
    if (img?.data && img?.mimeType) {
      imageParts.push({
        inlineData: {
          mimeType: img.mimeType,
          data: img.data,
        },
      });
    }
  }

  if (!rawText.trim() && imageParts.length === 0) {
    return null;
  }

  return {
    sourceId,
    fileName: incoming.fileName || 'Branch Proposal',
    rawText,
    imageParts,
  };
}

function storeCachedSource(sourceId: string, entry: CachedBranchProposalSource): void {
  if (branchProposalCache.has(sourceId)) {
    branchProposalCache.delete(sourceId);
  }
  branchProposalCache.set(sourceId, entry);
  while (branchProposalCache.size > MAX_CACHED_SOURCES) {
    const oldestKey = branchProposalCache.keys().next().value;
    if (oldestKey === undefined) break;
    branchProposalCache.delete(oldestKey);
  }
}

function resolveBranchProposalSource(
  sourceId?: string,
  incoming?: IncomingBranchProposalSource | null
):
  | { ok: true; source: CachedBranchProposalSource }
  | { ok: false; status: number; error: string; code?: string } {
  const effectiveId = (incoming?.sourceId || sourceId || '').trim();

  if (incoming) {
    const built = buildCachedSource(effectiveId || `bp_${Date.now()}`, incoming);
    if (!built) {
      return {
        ok: false,
        status: 400,
        error:
          'Branch Proposal is not uploaded or readable. Please upload the Branch Proposal first before running AI scrutiny.',
      };
    }
    if (effectiveId) {
      storeCachedSource(effectiveId, built);
    }
    return { ok: true, source: built };
  }

  if (effectiveId) {
    const hit = branchProposalCache.get(effectiveId);
    if (hit) {
      return { ok: true, source: hit };
    }
    return {
      ok: false,
      status: 404,
      code: 'SOURCE_CACHE_MISS',
      error: 'Cached Branch Proposal source not found on server; resending full source.',
    };
  }

  return {
    ok: false,
    status: 400,
    error:
      'Branch Proposal is not uploaded or readable. Please upload the Branch Proposal first before running AI scrutiny.',
  };
}

const SYNTHESIS_CONTEXT_SECTIONS = new Set<ScrutinySectionKey>([
  'justification',
  'sec8SourceOfRepayment',
  'clientPrayer',
  'sec13BranchRecommendation',
]);

function buildCompactSectionContext(
  section: ScrutinySectionKey,
  context?: Record<string, unknown>
): string {
  if (!context) return '{}';

  if (section === 'samdObservation') {
    const { branchProposal: _omitBp, ...rest } = context;
    return JSON.stringify(rest);
  }

  if (SYNTHESIS_CONTEXT_SECTIONS.has(section)) {
    return JSON.stringify({
      placeholders: context.placeholders ?? {},
      extractedSummaries: context.extractedSummaries ?? {},
    });
  }

  // Pure Branch Proposal extraction sections read directly from the uploaded Branch Proposal
  return '{}';
}

function sanitizeAiMemoText(raw: string): string {
  let cleaned = raw
    .replace(/\*\*(.*?)\*\*/g, '$1')
    .replace(/^\s*\*\s+/gm, '')
    .trim();

  // Remove unwanted conversational AI preambles or processing disclaimers
  const unwantedLinePrefix =
    /^(?:(?:here\s+is|below\s+is|following\s+is)\s+(?:the\s+)?(?:generated|extracted|formal|requested)[^:\n]*:?\s*|ai[-\s]*generated\s*(?:content|text|section|summary)\s*:?\s*|based\s+on\s+(?:a\s+thorough\s+)?scrutiny\s+of\s+the\s+uploaded\s+branch\s+proposal[,:\s]*)/i;
  cleaned = cleaned.replace(unwantedLinePrefix, '').trim();

  const metaDisclaimerLine =
    /^(?:\[?\s*(?:ai[-\s]*note|processing\s+note|disclaimer|internal\s+note|note\s*:\s*(?:based\s+on\s+the\s+uploaded|information\s+(?:regarding|is)\s+(?:unavailable|insufficient|not\s+provided)))[^\n]*)$/gim;
  cleaned = cleaned.replace(metaDisclaimerLine, '').replace(/\n{3,}/g, '\n\n').trim();

  // If the entire text is literally a "no info / unavailable" meta-message, return "" so clean memo body isn't polluted
  if (
    /^(?:information\s+(?:regarding|on|for)[^.]{0,80}\s+is\s+)?(?:unavailable|insufficient|not\s+(?:specified|provided|mentioned|available))(?:\s+in\s+the\s+(?:uploaded\s+)?branch\s+proposal)?\.?$/i.test(
      cleaned
    )
  ) {
    return '';
  }

  return cleaned;
}

const NARRATIVE_SECTION_KEYS = new Set<ScrutinySectionKey>([
  'justification',
  'sec1NatureOfBusiness',
  'sec8SourceOfRepayment',
  'clientPrayer',
  'branchVisit',
  'sec13BranchRecommendation',
  'sec14IccdRecommendation',
  'samdObservation',
]);

function evaluateSectionResult(
  section: ScrutinySectionKey,
  parsed: Record<string, unknown>
): { result: Record<string, unknown>; insufficient: boolean; insufficientReason: string } {
  const insufficientPattern =
    /\b(unavailable|insufficient|not\s+(?:specified|provided|mentioned|available|stated)|lacks?\s+sufficient|left\s+blank)\b/i;

  if (NARRATIVE_SECTION_KEYS.has(section)) {
    const rawText = String(parsed.text ?? '').trim();
    const cleanText = sanitizeAiMemoText(rawText);
    const modelReason = String(parsed.insufficientReason ?? '').trim();

    let reason = modelReason;
    if (!cleanText && !reason) {
      reason = insufficientPattern.test(rawText)
        ? rawText
        : 'Necessary details were not found in the uploaded Branch Proposal.';
    }

    const insufficient = Boolean(reason) || cleanText === '';
    return {
      result: { text: cleanText, insufficientReason: reason },
      insufficient,
      insufficientReason: reason,
    };
  }

  const summary = String(parsed.summary ?? '').trim();
  let reason = insufficientPattern.test(summary) ? summary : '';

  if (section === 'sec2LiabilityOurBank') {
    const hasAnyField = Boolean(
      String(parsed.limit ?? '').trim() ||
        String(parsed.validity ?? '').trim() ||
        String(parsed.remarks ?? '').trim() ||
        String(parsed.nonFundedNote ?? '').trim()
    );
    if (!hasAnyField && !reason) {
      reason = summary || 'Section 2 liability details were blank or unavailable in the Branch Proposal.';
    }
  } else if (section === 'sec10ClientParticulars') {
    const particulars = (parsed.particulars ?? {}) as Record<string, unknown>;
    const filledCount = Object.values(particulars).filter((v) => String(v ?? '').trim() !== '').length;
    if (filledCount === 0 && !reason) {
      reason = summary || 'Client particulars fields were blank or unavailable in the Branch Proposal.';
    }
  }

  return {
    result: parsed,
    insufficient: Boolean(reason),
    insufficientReason: reason,
  };
}

const FREE_TIER_MODELS = ['gemini-4-hydrogen', 'gemini-3.1-flash-lite'] as const;
type SupportedScrutinyModel = (typeof FREE_TIER_MODELS)[number];

const modelCooldownUntil = new Map<string, number>();

function isQuotaOrRateError(msg: string): boolean {
  return /\b429\b|RESOURCE_EXHAUSTED|quota|rate\s*limit/i.test(msg);
}

function isDailyQuotaError(msg: string): boolean {
  return /GenerateRequestsPerDay|per\s*day|daily|free_tier/i.test(msg);
}

function isModelNotFoundError(msg: string): boolean {
  return /\b404\b|NOT_FOUND|is not found|not supported for generateContent/i.test(msg);
}

function formatGeminiError(err: unknown, fallback = 'Failed to scrutinize section.'): string {
  if (err instanceof MissingApiKeyError) {
    return err.message;
  }
  const raw = err instanceof Error ? err.message : String(err ?? '');
  if (!raw) return fallback;
  try {
    const parsed = JSON.parse(raw) as { error?: { code?: number; status?: string; message?: string } };
    if (parsed?.error) {
      const code = parsed.error.code;
      const status = parsed.error.status ?? '';
      if (code === 429 || status === 'RESOURCE_EXHAUSTED') {
        return isDailyQuotaError(raw)
          ? 'Gemini API daily request quota reached across available free-tier models — click Retry once quota resets or select a billing-enabled API key in Settings > Secrets.'
          : 'Gemini API rate limit reached — click Retry in a moment to generate this section.';
      }
      if (
        code === 403 ||
        status === 'PERMISSION_DENIED' ||
        /API_KEY_INVALID|API key not valid/i.test(parsed.error.message ?? raw)
      ) {
        return 'Gemini API key is invalid or lacks permission. Please verify or select a valid API key in Settings > Secrets.';
      }
      if (parsed.error.message) {
        return parsed.error.message.split('\n')[0];
      }
    }
  } catch {
    // Plain text error message
  }
  if (/API_KEY_INVALID|API key not valid|PERMISSION_DENIED/i.test(raw)) {
    return 'Gemini API key is invalid or lacks permission. Please verify or select a valid API key in Settings > Secrets.';
  }
  if (isQuotaOrRateError(raw)) {
    return isDailyQuotaError(raw)
      ? 'Gemini API daily request quota reached across available free-tier models — click Retry once quota resets or select a billing-enabled API key in Settings > Secrets.'
      : 'Gemini API rate limit reached — click Retry in a moment to generate this section.';
  }
  return raw.split('\n')[0] || fallback;
}

function buildModelConfig(
  model: SupportedScrutinyModel,
  baseConfig: Record<string, unknown>
): Record<string, unknown> {
  if (model === 'gemini-4-hydrogen') {
    return {
      ...baseConfig,
      thinkingConfig: { thinkingLevel: ThinkingLevel.LOW },
    };
  }
  return baseConfig;
}

function stripJsonFences(raw: string): string {
  const trimmed = raw.trim();
  if (!trimmed.startsWith('```')) return trimmed;
  return trimmed
    .replace(/^```(?:json)?\s*/i, '')
    .replace(/\s*```$/i, '')
    .trim();
}

async function runWithConcurrency<T>(
  items: T[],
  limit: number,
  worker: (item: T) => Promise<void>
): Promise<void> {
  if (items.length === 0) return;
  let index = 0;
  const runners = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (index < items.length) {
      const current = items[index++];
      await worker(current);
    }
  });
  await Promise.all(runners);
}

async function executeWithModelFallback<T>(
  runner: (model: SupportedScrutinyModel) => Promise<T>,
  shouldStop?: () => boolean
): Promise<T> {
  const maxPasses = 4;
  let lastErr: unknown;

  for (let pass = 0; pass < maxPasses; pass++) {
    const now = Date.now();
    const active = FREE_TIER_MODELS.filter((m) => (modelCooldownUntil.get(m) ?? 0) <= now);
    const candidates: SupportedScrutinyModel[] =
      active.length > 0 ? active : [...FREE_TIER_MODELS];

    let sawTransientRateBurst = false;

    for (let idx = 0; idx < candidates.length; idx++) {
      if (shouldStop?.()) {
        throw new Error('Request cancelled.');
      }
      const model = candidates[idx];
      try {
        const result = await runner(model);
        modelCooldownUntil.delete(model);
        return result;
      } catch (err) {
        lastErr = err;
        const msg = err instanceof Error ? err.message : String(err);

        if (isDailyQuotaError(msg) || isModelNotFoundError(msg)) {
          modelCooldownUntil.set(model, Date.now() + 10 * 60_000);
          continue;
        }

        if (isQuotaOrRateError(msg)) {
          sawTransientRateBurst = true;
          modelCooldownUntil.set(model, Date.now() + 1_500);
          continue;
        }

        const isTransientNet = /\b(500|502|503|504)\b|timeout|network|econnreset/i.test(msg);
        if (isTransientNet && !shouldStop?.()) {
          await new Promise((resolve) => setTimeout(resolve, 250));
          try {
            return await runner(model);
          } catch (retryErr) {
            lastErr = retryErr;
            continue;
          }
        }

        throw err;
      }
    }

    if (sawTransientRateBurst && pass < maxPasses - 1 && !shouldStop?.()) {
      const baseDelay = 500 + pass * 550;
      const jitterMs = baseDelay + Math.floor(Math.random() * 300);
      await new Promise((resolve) => setTimeout(resolve, jitterMs));
      continue;
    }
    break;
  }

  throw lastErr;
}

const MAX_CACHED_RESULTS = 200;
const sectionResultCache = new Map<
  string,
  { result: Record<string, unknown>; insufficient: boolean; insufficientReason: string }
>();

function hashString(input: string): string {
  return createHash('sha256').update(input).digest('hex').slice(0, 16);
}

function getSectionCacheKey(
  sourceId: string,
  section: ScrutinySectionKey,
  compactContext: string
): string {
  return `${sourceId}::${section}::${hashString(compactContext)}`;
}

function storeSectionCache(
  key: string,
  value: { result: Record<string, unknown>; insufficient: boolean; insufficientReason: string }
): void {
  if (sectionResultCache.has(key)) {
    sectionResultCache.delete(key);
  }
  sectionResultCache.set(key, value);
  while (sectionResultCache.size > MAX_CACHED_RESULTS) {
    const oldest = sectionResultCache.keys().next().value;
    if (oldest === undefined) break;
    sectionResultCache.delete(oldest);
  }
}

const SCRUTINY_SYSTEM_INSTRUCTION =
  'You are a Senior Credit Analyst at the Special Asset Management Division (SAMD), Head Office, Sammilito Islami Bank PLC (Union Bank PLC Unit), Dhaka. ' +
  'Your primary source document is the uploaded Branch Proposal (both its page images and extracted text). ' +
  'Scrutinize the Branch Proposal thoroughly and produce content strictly tailored to the requested IC Memo section. ' +
  'CRITICAL INTEGRITY RULE: Never invent, assume, or fabricate any facts, figures, dates, case numbers, valuations, names, observations, or recommendations not supported by the Branch Proposal (and verified memo figures provided in context). ' +
  'OUTPUT PURITY RULE: Output ONLY the official formal banking memorandum text/fields matching the section heading. NEVER add conversational remarks ("Here is...", "Based on the uploaded document..."), disclaimers, AI processing notes, scrutiny commentary, or markdown asterisks inside memo text. If required information is unavailable in the Branch Proposal, leave missing fields blank ("") and report what is missing strictly inside "insufficientReason" (or "summary" for structured tables).';

async function scrutinizeSingleSection(
  ai: GoogleGenAI,
  section: ScrutinySectionKey,
  source: CachedBranchProposalSource,
  context?: Record<string, unknown>,
  shouldStop?: () => boolean,
  bypassCache = false
): Promise<{ result: Record<string, unknown>; insufficient: boolean; insufficientReason: string }> {
  const { prompt, schema } = getScrutinyConfig(section);
  const compactContext = buildCompactSectionContext(section, context);
  const cacheKey = getSectionCacheKey(source.sourceId, section, compactContext);

  if (!bypassCache) {
    const cached = sectionResultCache.get(cacheKey);
    if (cached) {
      return cached;
    }
  }

  const parts: Array<
    | { inlineData: { mimeType: string; data: string } }
    | { text: string }
  > = [
    ...source.imageParts,
    {
      text: `${prompt}

Uploaded Branch Proposal File: ${source.fileName}

Extracted Branch Proposal Text (may contain minor scanner OCR noise — cross-verify against attached page images when present):
${source.rawText}

Additional Verified IC Memo Context (from other uploaded documents & manual inputs):
${compactContext}`,
    },
  ];

  const response = await executeWithModelFallback(
    (model) =>
      ai.models.generateContent({
        model,
        contents: { parts },
        config: buildModelConfig(model, {
          systemInstruction: SCRUTINY_SYSTEM_INSTRUCTION,
          responseMimeType: 'application/json',
          responseSchema: schema,
        }),
      }),
    shouldStop
  );

  const rawJson = stripJsonFences(response.text ?? '{}');
  const parsed = JSON.parse(rawJson || '{}') as Record<string, unknown>;
  const evaluated = evaluateSectionResult(section, parsed);
  storeSectionCache(cacheKey, evaluated);
  return evaluated;
}

/**
 * Incrementally extracts completed top-level `"key": { ... }` object entries from a streaming JSON root object.
 */
function extractCompletedTopLevelObjects(
  jsonBuffer: string,
  emittedKeys: Set<string>
): Array<{ key: string; value: Record<string, unknown> }> {
  const out: Array<{ key: string; value: Record<string, unknown> }> = [];
  const len = jsonBuffer.length;
  let i = 0;

  while (i < len && /\s/.test(jsonBuffer[i])) i++;
  if (i >= len || jsonBuffer[i] !== '{') return out;
  i++;

  while (i < len) {
    while (i < len && (/\s/.test(jsonBuffer[i]) || jsonBuffer[i] === ',')) i++;
    if (i >= len || jsonBuffer[i] === '}') break;

    if (jsonBuffer[i] !== '"') break;
    const keyStart = i;
    i++;
    let escaped = false;
    let keyClosed = false;
    while (i < len) {
      const ch = jsonBuffer[i];
      if (escaped) {
        escaped = false;
      } else if (ch === '\\') {
        escaped = true;
      } else if (ch === '"') {
        keyClosed = true;
        i++;
        break;
      }
      i++;
    }
    if (!keyClosed) break;

    let key: string;
    try {
      key = JSON.parse(jsonBuffer.slice(keyStart, i)) as string;
    } catch {
      break;
    }

    while (i < len && /\s/.test(jsonBuffer[i])) i++;
    if (i >= len || jsonBuffer[i] !== ':') break;
    i++;
    while (i < len && /\s/.test(jsonBuffer[i])) i++;
    if (i >= len) break;

    if (jsonBuffer[i] !== '{') {
      break;
    }

    const valStart = i;
    let depth = 0;
    let inStr = false;
    let strEsc = false;
    let valEnd = -1;

    while (i < len) {
      const ch = jsonBuffer[i];
      if (inStr) {
        if (strEsc) {
          strEsc = false;
        } else if (ch === '\\') {
          strEsc = true;
        } else if (ch === '"') {
          inStr = false;
        }
      } else {
        if (ch === '"') {
          inStr = true;
        } else if (ch === '{') {
          depth++;
        } else if (ch === '}') {
          depth--;
          if (depth === 0) {
            valEnd = i;
            i++;
            break;
          }
        }
      }
      i++;
    }

    if (valEnd === -1) {
      break;
    }

    if (!emittedKeys.has(key)) {
      try {
        const parsedObj = JSON.parse(jsonBuffer.slice(valStart, valEnd + 1)) as Record<
          string,
          unknown
        >;
        emittedKeys.add(key);
        out.push({ key, value: parsedObj });
      } catch {
        // Defer to full JSON parse at stream end
      }
    }
  }

  return out;
}

function buildConsolidatedBatchSpec(sections: ScrutinySectionKey[]): {
  prompt: string;
  schema: Schema;
} {
  const properties: Record<string, Schema> = {};
  const promptBlocks: string[] = [];

  for (const sec of sections) {
    const cfg = getScrutinyConfig(sec);
    properties[sec] = cfg.schema;
    promptBlocks.push(`### SECTION KEY "${sec}":\n${cfg.prompt}`);
  }

  const prompt = `Scrutinize the uploaded Branch Proposal thoroughly once and populate ALL ${sections.length} requested IC Memo section keys below in a single structured JSON object.
Follow EVERY section's exact instructions and integrity rules:
- Never fabricate, assume, or invent any facts, dates, figures, or observations.
- Output ONLY clean formal banking text/data for each section (no conversational preambles, AI notes, or markdown asterisks).
- If information for a specific section is absent or insufficient in the Branch Proposal, keep missing fields blank ("") and state why strictly inside that section's "insufficientReason" or "summary" field.

${promptBlocks.join('\n\n')}`;

  return {
    prompt,
    schema: {
      type: Type.OBJECT,
      properties,
      required: sections,
      propertyOrdering: sections,
    },
  };
}

async function startServer() {
  const app = express();
  app.use(express.json({ limit: '50mb' }));

  // Pre-warm server-side Branch Proposal source cache so subsequent section/batch calls skip re-uploading base64 images
  app.post('/api/gemini/cache-source', (req, res) => {
    try {
      const { sourceId, branchProposalSource } = req.body as {
        sourceId?: string;
        branchProposalSource?: IncomingBranchProposalSource | null;
      };
      const resolved = resolveBranchProposalSource(sourceId, branchProposalSource);
      if (!resolved.ok) {
        res.status(resolved.status).json({ error: resolved.error, code: resolved.code });
        return;
      }
      res.json({ cached: true, sourceId: resolved.source.sourceId });
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Failed to cache source.';
      res.status(500).json({ error: message });
    }
  });

  // Multimodal Branch Proposal Scrutiny for any individual IC Memo section (Justification .. Section 15)
  app.post('/api/gemini/scrutinize-section', async (req, res) => {
    try {
      const { section, sourceId, branchProposalSource, context, bypassCache } = req.body as {
        section: ScrutinySectionKey;
        sourceId?: string;
        branchProposalSource?: IncomingBranchProposalSource | null;
        context?: Record<string, unknown>;
        bypassCache?: boolean;
      };

      if (!section || !VALID_SCRUTINY_SECTIONS.has(section)) {
        res.status(400).json({ error: 'Invalid or missing section key.' });
        return;
      }

      const resolved = resolveBranchProposalSource(sourceId, branchProposalSource);
      if (!resolved.ok) {
        res.status(resolved.status).json({ error: resolved.error, code: resolved.code });
        return;
      }

      const aiClient = resolveAiClient();
      if (!aiClient.ok) {
        res.status(aiClient.status).json({ error: aiClient.error });
        return;
      }

      const { result, insufficient, insufficientReason } = await scrutinizeSingleSection(
        aiClient.ai,
        section,
        resolved.source,
        context,
        undefined,
        Boolean(bypassCache)
      );
      res.json({ result, insufficient, insufficientReason, sourceId: resolved.source.sourceId });
    } catch (error) {
      if (!(error instanceof MissingApiKeyError)) {
        console.error('Error scrutinizing IC Memo section:', error);
      }
      const status = error instanceof MissingApiKeyError ? 503 : 500;
      const message = formatGeminiError(error, 'Failed to scrutinize section.');
      res.status(status).json({ error: message });
    }
  });

  // Consolidated Streaming ("Generate and Fill Up All Together") via SSE:
  // Executes all independent sections in 1 consolidated streaming Gemini call (emitting each section incrementally as its JSON property finishes)
  // followed by 1 SAMD synthesis call — reducing quota consumption by 8x-15x while preserving live progressive UI updates.
  app.post('/api/gemini/scrutinize-all', async (req, res) => {
    try {
      const { sections, sourceId, branchProposalSource, context } = req.body as {
        sections?: ScrutinySectionKey[];
        sourceId?: string;
        branchProposalSource?: IncomingBranchProposalSource | null;
        context?: Record<string, unknown>;
      };

      const targetSections: ScrutinySectionKey[] = (
        Array.isArray(sections) && sections.length > 0
          ? sections
          : Array.from(VALID_SCRUTINY_SECTIONS)
      ).filter((s): s is ScrutinySectionKey => VALID_SCRUTINY_SECTIONS.has(s));

      if (targetSections.length === 0) {
        res.status(400).json({ error: 'No valid sections specified for batch scrutiny.' });
        return;
      }

      const resolved = resolveBranchProposalSource(sourceId, branchProposalSource);
      if (!resolved.ok) {
        res.status(resolved.status).json({ error: resolved.error, code: resolved.code });
        return;
      }

      const aiClient = resolveAiClient();
      if (!aiClient.ok) {
        res.status(aiClient.status).json({ error: aiClient.error });
        return;
      }
      const ai = aiClient.ai;

      res.setHeader('Content-Type', 'text/event-stream; charset=utf-8');
      res.setHeader('Cache-Control', 'no-cache, no-transform');
      res.setHeader('Connection', 'keep-alive');
      res.setHeader('X-Accel-Buffering', 'no');
      res.flushHeaders?.();

      let clientClosed = false;
      res.on('close', () => {
        if (!res.writableEnded) {
          clientClosed = true;
        }
      });

      const writeEvent = (payload: Record<string, unknown>) => {
        if (clientClosed || res.writableEnded) return;
        res.write(`data: ${JSON.stringify(payload)}\n\n`);
      };

      const includesSamd = targetSections.includes('samdObservation');
      const independentQueue = targetSections.filter((s) => s !== 'samdObservation');
      const samdContext: Record<string, unknown> = { ...(context ?? {}) };

      const feedSamdContext = (section: ScrutinySectionKey, result: Record<string, unknown>) => {
        if (section === 'justification' && result.text) {
          samdContext.justificationLines = String(result.text);
        } else if (section === 'sec8SourceOfRepayment' && result.text) {
          samdContext.sourceOfRepayment = String(result.text);
        } else if (section === 'clientPrayer' && result.text) {
          samdContext.clientPrayer = String(result.text);
        } else if (section === 'branchVisit' && result.text) {
          samdContext.branchVisitReport = String(result.text);
        } else if (section === 'sec13BranchRecommendation' && result.text) {
          samdContext.sec13BranchRecommendation = String(result.text);
        } else if (section === 'sec14IccdRecommendation' && result.text) {
          samdContext.sec14IccdRecommendation = String(result.text);
        }
      };

      const emitEvaluatedSection = (
        section: ScrutinySectionKey,
        evaluated: {
          result: Record<string, unknown>;
          insufficient: boolean;
          insufficientReason: string;
        }
      ) => {
        feedSamdContext(section, evaluated.result);
        writeEvent({
          type: 'section',
          section,
          result: evaluated.result,
          insufficient: evaluated.insufficient,
          insufficientReason: evaluated.insufficientReason,
        });
      };

      // Step 1: Serve any already-cached unchanged sections immediately
      const uncachedQueue: ScrutinySectionKey[] = [];
      for (const sec of independentQueue) {
        const compactCtx = buildCompactSectionContext(sec, context);
        const cacheKey = getSectionCacheKey(resolved.source.sourceId, sec, compactCtx);
        const cached = sectionResultCache.get(cacheKey);
        if (cached) {
          emitEvaluatedSection(sec, cached);
        } else {
          uncachedQueue.push(sec);
        }
      }

      // Step 2: Partition uncached sections into 4 balanced streaming groups (+ concurrent SAMD)
      // Halves multimodal image upload pressure & prevents RPM/TPM burst 429s while streaming every section live.
      const PARALLEL_GROUPS: ScrutinySectionKey[][] = [
        ['justification', 'sec1NatureOfBusiness', 'clientPrayer', 'sec8SourceOfRepayment'],
        ['branchVisit', 'sec13BranchRecommendation', 'sec14IccdRecommendation'],
        ['sec2LiabilityOurBank', 'sec3LiabilityOtherUnits', 'sec4CibStatus', 'sec5LegalAction', 'sec6AuditCompliance'],
        ['sec7Security', 'sec9HoSanctions', 'sec10ClientParticulars', 'sec11Financials'],
      ];

      const assigned = new Set<ScrutinySectionKey>(PARALLEL_GROUPS.flat());
      const activeGroups = PARALLEL_GROUPS.map((g) =>
        g.filter((sec) => uncachedQueue.includes(sec))
      ).filter((g) => g.length > 0);
      const extraSections = uncachedQueue.filter((sec) => !assigned.has(sec));
      if (extraSections.length > 0) {
        activeGroups.push(extraSections);
      }

      const batchContext = JSON.stringify({
        placeholders: context?.placeholders ?? {},
        extractedSummaries: context?.extractedSummaries ?? {},
      });

      const runStreamingGroup = async (groupSections: ScrutinySectionKey[]): Promise<void> => {
        if (groupSections.length === 0 || clientClosed) return;

        if (groupSections.length === 1) {
          const sec = groupSections[0];
          try {
            const evaluated = await scrutinizeSingleSection(
              ai,
              sec,
              resolved.source,
              context,
              () => clientClosed
            );
            emitEvaluatedSection(sec, evaluated);
          } catch (err) {
            if (!clientClosed) {
              writeEvent({
                type: 'error',
                section: sec,
                error: formatGeminiError(err, 'Failed to scrutinize section.'),
              });
            }
          }
          return;
        }

        const emittedKeys = new Set<string>();
        let batchErrorMessage: string | null = null;
        let hitHardQuota = false;
        const needsMemoContext = groupSections.some((s) => SYNTHESIS_CONTEXT_SECTIONS.has(s));

        try {
          const { prompt, schema } = buildConsolidatedBatchSpec(groupSections);
          const contextSuffix = needsMemoContext
            ? `\n\nAdditional Verified IC Memo Context (from other uploaded documents & manual inputs):\n${batchContext}`
            : '';
          const parts: Array<
            | { inlineData: { mimeType: string; data: string } }
            | { text: string }
          > = [
            ...resolved.source.imageParts,
            {
              text: `${prompt}

Uploaded Branch Proposal File: ${resolved.source.fileName}

Extracted Branch Proposal Text (may contain minor scanner OCR noise — cross-verify against attached page images when present):
${resolved.source.rawText}${contextSuffix}`,
            },
          ];

          const fullBuffer = await executeWithModelFallback(async (model) => {
            const stream = await ai.models.generateContentStream({
              model,
              contents: { parts },
              config: buildModelConfig(model, {
                systemInstruction: SCRUTINY_SYSTEM_INSTRUCTION,
                responseMimeType: 'application/json',
                responseSchema: schema,
              }),
            });

            let acc = '';
            for await (const chunk of stream) {
              if (clientClosed) break;
              const delta = chunk.text ?? '';
              if (!delta) continue;
              acc += delta;

              const completed = extractCompletedTopLevelObjects(acc, emittedKeys);
              for (const item of completed) {
                const sec = item.key as ScrutinySectionKey;
                if (!groupSections.includes(sec)) continue;
                const evaluated = evaluateSectionResult(sec, item.value);
                const cacheKey = getSectionCacheKey(
                  resolved.source.sourceId,
                  sec,
                  buildCompactSectionContext(sec, context)
                );
                storeSectionCache(cacheKey, evaluated);
                emitEvaluatedSection(sec, evaluated);
              }
            }
            return acc;
          }, () => clientClosed);

          if (!clientClosed && fullBuffer.trim()) {
            const cleanedBuffer = stripJsonFences(fullBuffer);
            const salvaged = extractCompletedTopLevelObjects(cleanedBuffer, emittedKeys);
            for (const item of salvaged) {
              const sec = item.key as ScrutinySectionKey;
              if (!groupSections.includes(sec)) continue;
              const evaluated = evaluateSectionResult(sec, item.value);
              const cacheKey = getSectionCacheKey(
                resolved.source.sourceId,
                sec,
                buildCompactSectionContext(sec, context)
              );
              storeSectionCache(cacheKey, evaluated);
              emitEvaluatedSection(sec, evaluated);
            }
            try {
              const finalParsed = JSON.parse(cleanedBuffer) as Record<
                string,
                Record<string, unknown>
              >;
              for (const sec of groupSections) {
                if (emittedKeys.has(sec)) continue;
                const rawObj = finalParsed[sec];
                if (rawObj && typeof rawObj === 'object') {
                  emittedKeys.add(sec);
                  const evaluated = evaluateSectionResult(sec, rawObj);
                  const cacheKey = getSectionCacheKey(
                    resolved.source.sourceId,
                    sec,
                    buildCompactSectionContext(sec, context)
                  );
                  storeSectionCache(cacheKey, evaluated);
                  emitEvaluatedSection(sec, evaluated);
                }
              }
            } catch {
              // Handled below by missing-section recovery
            }
          }
        } catch (err) {
          const rawMsg = err instanceof Error ? err.message : String(err);
          hitHardQuota = isQuotaOrRateError(rawMsg) && isDailyQuotaError(rawMsg);
          batchErrorMessage = formatGeminiError(err, 'Failed to scrutinize section.');
        }

        const remaining = groupSections.filter((sec) => !emittedKeys.has(sec));
        if (remaining.length > 0 && !clientClosed) {
          if (hitHardQuota && batchErrorMessage) {
            for (const sec of remaining) {
              writeEvent({
                type: 'error',
                section: sec,
                error: batchErrorMessage,
              });
            }
          } else {
            await runWithConcurrency(remaining, 2, async (sec) => {
              if (clientClosed) return;
              try {
                const evaluated = await scrutinizeSingleSection(
                  ai,
                  sec,
                  resolved.source,
                  context,
                  () => clientClosed
                );
                emitEvaluatedSection(sec, evaluated);
              } catch (secErr) {
                if (!clientClosed) {
                  writeEvent({
                    type: 'error',
                    section: sec,
                    error: formatGeminiError(
                      secErr,
                      batchErrorMessage || 'Failed to scrutinize section.'
                    ),
                  });
                }
              }
            });
          }
        }
      };

      const runSamdObservation = async (): Promise<void> => {
        if (!includesSamd || clientClosed) return;
        const samdCompactCtx = buildCompactSectionContext('samdObservation', samdContext);
        const samdCacheKey = getSectionCacheKey(
          resolved.source.sourceId,
          'samdObservation',
          samdCompactCtx
        );
        const cachedSamd = sectionResultCache.get(samdCacheKey);
        if (cachedSamd) {
          emitEvaluatedSection('samdObservation', cachedSamd);
          return;
        }
        try {
          const evaluated = await scrutinizeSingleSection(
            ai,
            'samdObservation',
            resolved.source,
            samdContext,
            () => clientClosed
          );
          emitEvaluatedSection('samdObservation', evaluated);
        } catch (err) {
          if (!clientClosed) {
            writeEvent({
              type: 'error',
              section: 'samdObservation',
              error: formatGeminiError(err, 'Failed to scrutinize SAMD section.'),
            });
          }
        }
      };

      await Promise.all([
        ...activeGroups.map((group) => runStreamingGroup(group)),
        runSamdObservation(),
      ]);

      if (!clientClosed) {
        writeEvent({ type: 'done', sourceId: resolved.source.sourceId });
        res.end();
      }
    } catch (error) {
      if (!(error instanceof MissingApiKeyError)) {
        console.error('Error during batch scrutiny:', error);
      }
      const status = error instanceof MissingApiKeyError ? 503 : 500;
      const message = formatGeminiError(error, 'Batch scrutiny failed.');
      if (!res.headersSent) {
        res.status(status).json({ error: message });
      } else if (!res.writableEnded) {
        res.write(
          `data: ${JSON.stringify({
            type: 'fatal',
            error: message,
          })}\n\n`
        );
        res.end();
      }
    }
  });

  const VALID_STAGE_DOCS = new Set<StageMemoDocType>(['mec', 'board', 'sanction']);
  const STAGE_SECTIONS_MAP: Record<StageMemoDocType, StageScrutinySectionKey[]> = {
    mec: ['mecIcRecommendation'],
    board: ['boardMecRecommendation'],
    sanction: ['sanctionCrossVerifiedSummary'],
  };

  function getStageSectionConfig(
    section: StageScrutinySectionKey,
    context?: Record<string, unknown>
  ): { prompt: string; schema: Schema } {
    const stageInputs = (context?.stageInputs ?? {}) as Record<string, string>;
    const icNo = (stageInputs.icMemoNumber || '').trim() || '[Chronologic Order of the IC Memo]';
    const icDt = (stageInputs.icMemoDate || '').trim() || '[IC Memo Date]';
    const mecNo = (stageInputs.mecMemoNumber || '').trim() || '[Chronologic Order of the MEC Memo]';
    const mecDt = (stageInputs.mecMemoDate || '').trim() || '[MEC Memo Date]';
    const boardNo = (stageInputs.boardMemoNumber || '').trim() || '[Chronologic Order of the Board Memo]';
    const boardDt = (stageInputs.boardMemoDate || '').trim() || '[Board Memo Date]';

    const icCentralRule = `PRIMARY SOURCE INSTRUCTION: Treat the finalized IC Memo data in the context ("icMemoFinalized", "placeholders", "icMemoAiSections", and carried sections) as the primary central data reference. Preserve 100% consistency with every client detail, facility mode, financial figure, security value, justification, observation, and recommendation established in the IC Memo — never alter, contradict, omit, or invent inconsistent figures. Adapt only the framing/terminology for this approval stage.`;

    const schema: Schema = {
      type: Type.OBJECT,
      properties: {
        text: strField('Clean official formal banking text for this memo section, or "" if source inputs are absent.'),
        insufficientReason: strField('Empty string "" if supported; otherwise concise UI note listing missing fields/inputs.'),
      },
      required: ['text', 'insufficientReason'],
    };

    switch (section) {
      case 'mecIcRecommendation': {
        const icOrderToken = /\/26$/i.test(icNo) ? icNo : `${icNo}/26`;
        return {
          prompt: `${icCentralRule}
Write Section "16. Recommendation of the Investment Committee (IC)" for the Management Executive Committee (MEC) Memorandum.
You MUST use the EXACT standardized wording below, inserting Chronologic Order "${icOrderToken}" and IC Memo Date "${icDt}":
"This proposal was placed before the Investment Committee (IC) in its ${icOrderToken} meeting held on ${icDt}. The Committee recommended the proposal as presented."
Output ONLY that clean official sentence with no AI preambles, no markdown asterisks, and no extra commentary. If either Chronologic Order of the IC Memo or IC Memo Date is still an unfilled bracketed placeholder, report that concisely in "insufficientReason".`,
          schema,
        };
      }
      case 'boardMecRecommendation': {
        const mecOrderToken = /\/26$/i.test(mecNo) ? mecNo : `${mecNo}/26`;
        return {
          prompt: `${icCentralRule}
Write Section "17. Management Executive Committee (MEC) Recommendation" for the Board of Directors Memorandum.
You MUST use the EXACT standardized wording below, inserting Chronologic Order "${mecOrderToken}" and MEC Memo Date "${mecDt}":
"This proposal was placed before the Management Executive Committee (MEC) in its ${mecOrderToken} meeting held on ${mecDt}. The Committee recommended the proposal as presented."
Output ONLY that clean official sentence with no AI preambles, no markdown asterisks, and no extra commentary. If either Chronologic Order of the MEC Memo or MEC Memo Date is still an unfilled bracketed placeholder, report that concisely in "insufficientReason".`,
          schema,
        };
      }
      case 'sanctionCrossVerifiedSummary':
        return {
          prompt: `${icCentralRule}
Perform an AI cross-verification across ALL uploaded source documents (Outstanding, Party Performance, Overdue, Account Statement, CIB Report, Cost of Fund, Branch Proposal) and the chained IC -> MEC -> Board memo data to write Section "AI Cross-Verified Facility & Liability Confirmation" for the Sanction Letter.
Confirm exact verified figures in Million BDT (NET Outstanding #25#, Principal #7#, Charges #8#, Down Payment #35# in A/C #36#, Cost of Fund #31#, Breakup Total #32#, Proposed Waiver #4#, Settlement/Exit Amount #27#, Settlement Period #28#, Security MV #33# & FSV #34#) and confirm full consistency with the IC Memo, Branch Proposal, and Board sanction parameters. Clean official prose only; flag any missing verification input strictly inside "insufficientReason".`,
          schema,
        };
    }
  }

  function resolveOptionalStageSource(
    sourceId?: string,
    incoming?: IncomingBranchProposalSource | null
  ): CachedBranchProposalSource | null {
    if (incoming) {
      const id = (incoming.sourceId || sourceId || `stage_${Date.now()}`).trim();
      const built = buildCachedSource(id, incoming);
      if (built && (incoming.sourceId || sourceId)) {
        storeCachedSource(id, built);
      }
      return built;
    }
    if (sourceId && branchProposalCache.has(sourceId.trim())) {
      return branchProposalCache.get(sourceId.trim()) ?? null;
    }
    return null;
  }

  function buildCompactStageContext(context?: Record<string, unknown>): string {
    if (!context) return '{}';
    const { branchProposal: _omitTopBp, icMemoFinalized, ...rest } = context;
    if (icMemoFinalized && typeof icMemoFinalized === 'object') {
      const { branchProposal: _omitNestedBp, ...finalizedClean } = icMemoFinalized as Record<
        string,
        unknown
      >;
      return JSON.stringify({ ...rest, icMemoFinalized: finalizedClean });
    }
    return JSON.stringify(rest);
  }

  async function scrutinizeSingleStageSection(
    ai: GoogleGenAI,
    section: StageScrutinySectionKey,
    source: CachedBranchProposalSource | null,
    context?: Record<string, unknown>,
    shouldStop?: () => boolean,
    precomputedContextStr?: string
  ): Promise<{ result: Record<string, unknown>; insufficient: boolean; insufficientReason: string }> {
    const { prompt, schema } = getStageSectionConfig(section, context);
    const contextStr = precomputedContextStr ?? buildCompactStageContext(context);
    const cacheKey = `${source?.sourceId ?? 'ctx_only'}::STAGE::${section}::${hashString(contextStr)}`;
    const cached = sectionResultCache.get(cacheKey);
    if (cached) return cached;

    const textPayload = `${prompt}

${source ? `Uploaded Source Document (${source.fileName}):\n${source.rawText}\n` : ''}
Chained Memo & Verified Document Context (IC Memo -> MEC Memo -> Board Memo -> Sanction + Stage Inputs):
${contextStr}`;

    const icAi = (context?.icMemoAiSections ?? {}) as Record<string, unknown>;
    const hasPopulatedIc = Boolean(icAi.justification || icAi.samdObservation);
    const shouldAttachImages =
      Boolean(source) &&
      (section === 'sanctionCrossVerifiedSummary' || !hasPopulatedIc || (source?.rawText?.trim().length ?? 0) < 500);

    const parts: Array<
      | { inlineData: { mimeType: string; data: string } }
      | { text: string }
    > = shouldAttachImages && source
      ? [...source.imageParts, { text: textPayload }]
      : [{ text: textPayload }];

    const response = await executeWithModelFallback(
      (model) =>
        ai.models.generateContent({
          model,
          contents: { parts },
          config: buildModelConfig(model, {
            systemInstruction: SCRUTINY_SYSTEM_INSTRUCTION,
            responseMimeType: 'application/json',
            responseSchema: schema,
          }),
        }),
      shouldStop
    );

    const parsed = JSON.parse(stripJsonFences(response.text ?? '{}') || '{}') as Record<string, unknown>;
    const rawText = String(parsed.text ?? '').trim();
    const cleanText = sanitizeAiMemoText(rawText);
    const modelReason = String(parsed.insufficientReason ?? '').trim();
    const insufficient = Boolean(modelReason) || cleanText === '';
    const evaluated = {
      result: { text: cleanText, insufficientReason: modelReason },
      insufficient,
      insufficientReason:
        modelReason || (!cleanText ? 'Required chained memo context or source data is unavailable.' : ''),
    };
    storeSectionCache(cacheKey, evaluated);
    return evaluated;
  }

  function buildConsolidatedStageBatchSpec(
    sections: StageScrutinySectionKey[],
    context?: Record<string, unknown>
  ): { prompt: string; schema: Schema } {
    const properties: Record<string, Schema> = {};
    const promptBlocks: string[] = [];
    for (const sec of sections) {
      const cfg = getStageSectionConfig(sec, context);
      properties[sec] = cfg.schema;
      promptBlocks.push(`### STAGE SECTION KEY "${sec}":\n${cfg.prompt}`);
    }
    const prompt = `Scrutinize the finalized IC Memo, chained memo context, and any uploaded source memo once and populate ALL ${sections.length} requested stage sections below in a single structured JSON object.
Follow EVERY section's exact instructions and integrity rules cleanly (no AI notes or markdown asterisks).

${promptBlocks.join('\n\n')}`;

    return {
      prompt,
      schema: {
        type: Type.OBJECT,
        properties,
        required: sections,
        propertyOrdering: sections,
      },
    };
  }

  app.post('/api/gemini/extract-stage-source', async (req, res) => {
    try {
      const { targetDocType, sourceId, memoSource, stageInputs, existingContext } = req.body as {
        targetDocType: StageMemoDocType;
        sourceId?: string;
        memoSource?: IncomingBranchProposalSource | null;
        stageInputs?: Record<string, string>;
        existingContext?: Record<string, unknown>;
      };

      if (!targetDocType || !VALID_STAGE_DOCS.has(targetDocType)) {
        res.status(400).json({ error: 'Invalid target memo stage type.' });
        return;
      }

      const source = resolveOptionalStageSource(sourceId, memoSource);
      if (!source) {
        res.status(400).json({
          error: 'Uploaded document is empty or unreadable. Please upload a valid PDF, Word (.docx), Excel, or image memo file.',
        });
        return;
      }

      const aiClient = resolveAiClient();
      if (!aiClient.ok) {
        res.status(aiClient.status).json({ error: aiClient.error });
        return;
      }
      const ai = aiClient.ai;

      const sourceRole =
        targetDocType === 'mec'
          ? 'completed IC Memo'
          : targetDocType === 'board'
            ? 'completed MEC Memo'
            : 'completed Board Memo';

      const targetRole =
        targetDocType === 'mec'
          ? 'MEC Memo (Management Executive Committee)'
          : targetDocType === 'board'
            ? 'Board Memo (Board of Directors)'
            : 'Sanction Letter (Head Office Sanction Authority)';

      const kvSchema: Schema = {
        type: Type.OBJECT,
        properties: {
          key: strField('Field identifier'),
          value: strField('Exact extracted value or "" if absent'),
        },
        required: ['key', 'value'],
      };

      const extractSchema: Schema = {
        type: Type.OBJECT,
        properties: {
          placeholders: {
            type: Type.ARRAY,
            items: kvSchema,
            description:
              'List of {key, value} for placeholder codes found in the document from: #1# Proposal/Memo Date, #2# A/C Title & No, #3# Branch, #4# Proposed Waiver, #5# Mode of Investment, #6# Sanctioned Limit, #7# Principal Outstanding, #8# Charges, #10# Principal Balance for Settlement/COF, #11# Charged Profit, #12# Compensation, #25# NET Outstanding, #27# Settlement/Exit Amount, #28# Settlement Period, #31# Cost of Fund, #32# Breakup Total (10+8+11+12), #33# Security Market Value MV, #34# Security FSV, #35# Down Payment, #36# Down Payment Account, #42# Gross Balance for Max Waiver, #44# Max Permissible Waiver, #45# Surplus/Deficit vs COF, #47# Total Waiver Recommended, #51# Justification (3 points), #54# Nature of Business, #55# Classification Status (CL/DF/BL/SS), #56# First Sanction Date, #58# Expiry Date, #59# Uncharged Profit/Comp, #60# Last Repayment Date/Amt, #62# Provision Held, #64# Branch Proposal Ref No, #65# Branch Proposal Date, #68# Liability Cutoff Date, #69# CIS/Unique Customer ID.',
          },
          manualInputs: {
            type: Type.OBJECT,
            properties: {
              proposalDate: strField('Proposal / Memo Date (#1#), or ""'),
              waiverAmount: strField('Proposed Waiver (#4#) numeric/text without Tk prefix if clean, or ""'),
              settlementAmount: strField('Settlement / Exit Amount (#27#), or ""'),
              settlementPeriod: strField('Settlement Period (#28#), or ""'),
              waiverRecClientBranch: strField('Waiver Recommended by Branch/Client, or ""'),
              waiverRecHO: strField('Waiver Recommended by Head Office/SAMD, or ""'),
              sourceOfRepayment: strField('Section 8 Source of Repayment text, or ""'),
            },
            required: [
              'proposalDate',
              'waiverAmount',
              'settlementAmount',
              'settlementPeriod',
              'waiverRecClientBranch',
              'waiverRecHO',
              'sourceOfRepayment',
            ],
          },
          extractedStageInputs: {
            type: Type.OBJECT,
            properties: {
              icMemoNumber: strField('Chronologic Order / Number of IC Memo if explicitly stated in document, or ""'),
              icMemoDate: strField('IC Memo Date if explicitly stated, or ""'),
              mecMemoNumber: strField('Chronologic Order / Number of MEC Memo if explicitly stated, or ""'),
              mecMemoDate: strField('MEC Memo Date if explicitly stated, or ""'),
              boardMemoNumber: strField('Chronologic Order / Number of Board Memo if explicitly stated, or ""'),
              boardMemoDate: strField('Board Memo Date if explicitly stated, or ""'),
              sanctionRefNumber: strField('Sanction Reference No if explicitly stated, or ""'),
              sanctionDate: strField('Sanction Date if explicitly stated, or ""'),
            },
            required: [
              'icMemoNumber',
              'icMemoDate',
              'mecMemoNumber',
              'mecMemoDate',
              'boardMemoNumber',
              'boardMemoDate',
              'sanctionRefNumber',
              'sanctionDate',
            ],
          },
          narratives: {
            type: Type.OBJECT,
            properties: {
              justification: strField('Justification of Proposed Facility (#51#), exact points preserved or ""'),
              sec1NatureOfBusiness: strField('Section 1 Nature of Business (#54#), or ""'),
              sec8SourceOfRepayment: strField('Section 8 Source of Repayment, or ""'),
              clientPrayer: strField('Section 12 Contents of the Clients Prayer, or ""'),
              branchVisit: strField('Section 12 Branch Visit Report, or ""'),
              sec13BranchRecommendation: strField('Section 13 Recommendation of the Branch, or ""'),
              sec14IccdRecommendation: strField('Section 14 Recommendation of ICCD, or ""'),
              samdObservation: strField('Section 15 Observation and Recommendation of SAMD, or ""'),
              carriedMecIcRecommendation: strField('Recommendation of the IC (carried forward if present in uploaded MEC/Board Memo), or ""'),
              carriedBoardMecRecommendation: strField('Recommendation of the MEC (carried forward if present in uploaded Board Memo), or ""'),
            },
            required: [
              'justification',
              'sec1NatureOfBusiness',
              'sec8SourceOfRepayment',
              'clientPrayer',
              'branchVisit',
              'sec13BranchRecommendation',
              'sec14IccdRecommendation',
              'samdObservation',
              'carriedMecIcRecommendation',
              'carriedBoardMecRecommendation',
            ],
          },
          structuredTables: {
            type: Type.OBJECT,
            properties: {
              sec2Limit: strField('Section 2 Sanctioned Limit, or ""'),
              sec2Validity: strField('Section 2 Validity/Expiry, or ""'),
              sec2Remarks: strField('Section 2 Remarks, or ""'),
              sec2NonFundedNote: strField('Section 2 Non-Funded/sister concern note, or ""'),
              section3Rows: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    sl: strField('SL'),
                    nameOfConcern: strField('Name of Concern'),
                    bankOrNbfi: strField('Name of Banks/NBFIs'),
                    natureOfFacility: strField('Nature of Facility'),
                    limit: strField('Limit'),
                    outstanding: strField('Outstanding'),
                    overdueEol: strField('Overdue/EOL'),
                    expiryDate: strField('Expiry Date'),
                    clStatus: strField('CL Status'),
                  },
                  required: ['sl', 'nameOfConcern', 'bankOrNbfi', 'natureOfFacility', 'limit', 'outstanding', 'overdueEol', 'expiryDate', 'clStatus'],
                },
              },
              sec4CibDate: strField('Section 4 CIB Report Date, or ""'),
              section4Rows: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    status: strField('Status'),
                    unitBorrower: strField('Unit Borrower'),
                    unitGuarantor: strField('Unit Guarantor'),
                    otherBorrower: strField('Other Bank Borrower'),
                    otherGuarantor: strField('Other Bank Guarantor'),
                    remarks: strField('Remarks'),
                  },
                  required: ['status', 'unitBorrower', 'unitGuarantor', 'otherBorrower', 'otherGuarantor', 'remarks'],
                },
              },
              sec5StatusNote: strField('Section 5 Legal Action header note (e.g., Nil/N/A), or ""'),
              section5Rows: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    sl: strField('Sl'),
                    typeOfCase: strField('Type of Case'),
                    caseNo: strField('Case No.'),
                    filingDate: strField('Filing date'),
                    amount: strField('Amount'),
                    previousDate: strField('Previous date'),
                    nextDate: strField('Next date'),
                    presentStatus: strField('Present status'),
                  },
                  required: ['sl', 'typeOfCase', 'caseNo', 'filingDate', 'amount', 'previousDate', 'nextDate', 'presentStatus'],
                },
              },
              sec6StatusNote: strField('Section 6 Audit Status note, or ""'),
              section6Rows: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    sl: strField('SL'),
                    authority: strField('Authority'),
                    lastInspectionDate: strField('Last Inspection Date'),
                    objectionStatus: strField('Objection Status'),
                    objections: strField('Objections'),
                    branchReply: strField('Branch Reply'),
                  },
                  required: ['sl', 'authority', 'lastInspectionDate', 'objectionStatus', 'objections', 'branchReply'],
                },
              },
              sec7PrimaryNote: strField('Section 7a Primary note, or ""'),
              sec7PrimaryTotal: strField('Section 7a Primary Total, or ""'),
              sec7CollateralNote: strField('Section 7b Collateral note, or ""'),
              sec7CollateralSchedules: { type: Type.ARRAY, items: { type: Type.STRING } },
              sec7CoveragePct: strField('Section 7 Coverage %, or ""'),
              sec7OthersNote: strField('Section 7c Others note, or ""'),
              section9Rows: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    sanctionSl: strField('Sanction SL'),
                    sanctionAuthority: strField('Sanctioning Authority'),
                    approvedDate: strField('Approved Date'),
                    natureOfApproval: strField('Nature of Approval'),
                    complianceStatus: strField('Compliance Status'),
                  },
                  required: ['sanctionSl', 'sanctionAuthority', 'approvedDate', 'natureOfApproval', 'complianceStatus'],
                },
              },
              sec10ParticularsPairs: {
                type: Type.ARRAY,
                items: kvSchema,
                description:
                  'Key-value pairs from Section 10 Particulars using keys: nameOfClient, uniqueCustomerId, awcdMsdAccount, businessTypeNature, businessManagementType, clientType, keyPersonPhone, address, officeBusinessAddress, factoryAddress, purposeOfInvestment, limitOfInvestment, modeOfInvestment, lastVisitDate, stockPosition, fixedAsset, businessNetWorth, insuranceCoverage, tradeLicenseValidity, otherRegulatoryDocuments, dateOfOriginalInvestment, firstDisbursementManager, firstDateOfClassification.',
              },
              section10Directors: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    name: strField('Name of Director/Proprietor/Partner'),
                    age: strField('Age'),
                    education: strField('Educational Qualification'),
                    designation: strField('Designation'),
                    sharePct: strField('Share %'),
                    experience: strField('Business Experience'),
                    netWorth: strField('Personal Net Worth'),
                  },
                  required: ['name', 'age', 'education', 'designation', 'sharePct', 'experience', 'netWorth'],
                },
              },
              sec11FinancialsNote: strField('Section 11 Financials note, or ""'),
              section11Financials: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    year: strField('Year'),
                    grossSales: strField('Gross Sales'),
                    grossProfit: strField('Gross Profit'),
                    netProfit: strField('Net Profit'),
                    totalAsset: strField('Total Asset'),
                    totalLiabilities: strField('Total Liabilities'),
                    netWorth: strField('Net Worth'),
                  },
                  required: ['year', 'grossSales', 'grossProfit', 'netProfit', 'totalAsset', 'totalLiabilities', 'netWorth'],
                },
              },
              previousBoardDecisions: strField('Previous EC/Board decisions if stated, or ""'),
            },
            required: [
              'sec2Limit',
              'sec2Validity',
              'sec2Remarks',
              'sec2NonFundedNote',
              'section3Rows',
              'sec4CibDate',
              'section4Rows',
              'sec5StatusNote',
              'section5Rows',
              'sec6StatusNote',
              'section6Rows',
              'sec7PrimaryNote',
              'sec7PrimaryTotal',
              'sec7CollateralNote',
              'sec7CollateralSchedules',
              'sec7CoveragePct',
              'sec7OthersNote',
              'section9Rows',
              'sec10ParticularsPairs',
              'section10Directors',
              'sec11FinancialsNote',
              'section11Financials',
              'previousBoardDecisions',
            ],
          },
          targetStageSections: {
            type: Type.OBJECT,
            properties: {
              sanctionCrossVerifiedSummary: strField('For targetDocType sanction: AI Cross-Verified Facility & Liability Confirmation; else ""'),
            },
            required: [
              'sanctionCrossVerifiedSummary',
            ],
          },
          auditFindings: {
            type: Type.OBJECT,
            properties: {
              summary: strField('1-2 sentence audit summary of what was extracted from the uploaded memo.'),
              verifiedPoints: {
                type: Type.ARRAY,
                items: { type: Type.STRING },
                description: 'Key verified items extracted and cross-checked cleanly.',
              },
              discrepanciesOrMissing: {
                type: Type.ARRAY,
                items: { type: Type.STRING },
                description: 'Any missing fields, unreadable cells, arithmetic mismatches, or conflicting details detected in the uploaded memo (NEVER fabricate missing data).',
              },
            },
            required: ['summary', 'verifiedPoints', 'discrepanciesOrMissing'],
          },
        },
        required: [
          'placeholders',
          'manualInputs',
          'extractedStageInputs',
          'narratives',
          'structuredTables',
          'targetStageSections',
          'auditFindings',
        ],
      };

      const stageSpecificInstruction =
        targetDocType === 'mec'
          ? `TARGET STAGE: Prepare the MEC Memo from this uploaded completed IC Memo.
1. Extract ALL client particulars, account numbers (#2#, #69#), branch (#3#), dates (#1#, #68#), facility figures (#4#, #5#, #6#, #7#, #8#, #10#, #11#, #12#, #25#, #27#, #28#, #31#, #32#, #33#, #34#, #35#, #36#, #42#, #44#, #45#, #47#, #55#, #56#, #58#, #59#, #60#, #62#, #64#, #65#), structured tables (Sections 2-11), and all narratives (Justification #51#, Prayer, Branch Visit, Branch Rec, ICCD Rec, SAMD Observation) faithfully without omitting or altering figures.`
          : targetDocType === 'board'
            ? `TARGET STAGE: Prepare the Board Memo from this uploaded completed MEC Memo.
1. Extract ALL carried-forward client particulars, facility details, financial figures, tables (Sections 2-11), narratives (Sections 1, 8, 12-15), IC recommendation ("carriedMecIcRecommendation"), and any IC/MEC memo chronological orders & dates faithfully.`
            : `TARGET STAGE: Prepare the Head Office Sanction Letter from this uploaded completed Board Memo through deep AI scrutiny and cross-verification.
1. Extract ALL client particulars, facility details, approved financial figures, special terms & conditions, Board/MEC/IC references, and tables accurately.
2. Cross-verify arithmetic & textual consistency across all sections of the uploaded Board Memo (e.g., Principal + Charges + Charged Profit + Comp vs NET Outstanding; Breakup Total vs Settlement Amount & Cost of Fund; Waiver + Settlement vs Outstanding; Down Payment account & amount; Special Terms compliance). Report ANY discrepancy, missing figure, or conflict honestly in "auditFindings.discrepanciesOrMissing" without ever fabricating data!
3. Adapt and organize the approved data into formal Head Office Sanction Letter format inside "targetStageSections":
   - "sanctionCrossVerifiedSummary": AI Cross-Verified Facility & Liability Confirmation (exact Million BDT breakdown & cross-check outcome).`;

      const prompt = `You are a Senior Credit Analyst at SAMD, Head Office, Sammilito Islami Bank PLC.
Carefully read, extract, identify, and scrutinize EVERY page, table, and narrative from the uploaded ${sourceRole} ("${source.fileName}") to prepare the ${targetRole}.
NEVER omit, alter, or fabricate numbers, dates, account titles, or facts. Leave absent fields as "" and list any missing, unreadable, or inconsistent items inside "auditFindings.discrepanciesOrMissing".
Keep all numeric figures strictly in Million BDT matching the document (clean numbers like "1.61" where appropriate for placeholders).

Current User Stage Reference Inputs (use if document lacks them, or report if document differs):
${JSON.stringify(stageInputs ?? {})}

Existing Context Snapshot (if any prior docs were loaded):
${JSON.stringify(existingContext ? { placeholders: existingContext.placeholders } : {})}

${stageSpecificInstruction}

Uploaded ${sourceRole} Extracted Text:
${source.rawText}`;

      const parts: Array<{ inlineData: { mimeType: string; data: string } } | { text: string }> =
        source.imageParts.length > 0 ? [...source.imageParts, { text: prompt }] : [{ text: prompt }];

      const response = await executeWithModelFallback((model) =>
        ai.models.generateContent({
          model,
          contents: { parts },
          config: buildModelConfig(model, {
            systemInstruction: SCRUTINY_SYSTEM_INSTRUCTION,
            responseMimeType: 'application/json',
            responseSchema: extractSchema,
          }),
        })
      );

      const parsed = JSON.parse(stripJsonFences(response.text ?? '{}') || '{}') as Record<string, unknown>;
      res.json({
        sourceId: source.sourceId,
        fileName: source.fileName,
        extracted: parsed,
      });
    } catch (error) {
      if (!(error instanceof MissingApiKeyError)) {
        console.error('Error extracting stage source document:', error);
      }
      const status = error instanceof MissingApiKeyError ? 503 : 500;
      res.status(status).json({
        error: formatGeminiError(error, 'Failed to read and scrutinize uploaded memo document.'),
      });
    }
  });

  app.post('/api/gemini/scrutinize-stage-section', async (req, res) => {
    try {
      const { docType, section, sourceId, branchProposalSource, stageSource, context } = req.body as {
        docType: StageMemoDocType;
        section: StageScrutinySectionKey;
        sourceId?: string;
        branchProposalSource?: IncomingBranchProposalSource | null;
        stageSource?: IncomingBranchProposalSource | null;
        context?: Record<string, unknown>;
      };

      if (!docType || !VALID_STAGE_DOCS.has(docType) || !STAGE_SECTIONS_MAP[docType].includes(section)) {
        res.status(400).json({ error: 'Invalid memo stage or section key.' });
        return;
      }

      const source =
        resolveOptionalStageSource(sourceId, stageSource) ??
        resolveOptionalStageSource(sourceId, branchProposalSource);
      const aiClient = resolveAiClient();
      if (!aiClient.ok) {
        res.status(aiClient.status).json({ error: aiClient.error });
        return;
      }
      const evaluated = await scrutinizeSingleStageSection(aiClient.ai, section, source, context);
      res.json({
        result: evaluated.result,
        insufficient: evaluated.insufficient,
        insufficientReason: evaluated.insufficientReason,
        sourceId: source?.sourceId,
      });
    } catch (error) {
      if (!(error instanceof MissingApiKeyError)) {
        console.error('Error scrutinizing stage memo section:', error);
      }
      const status = error instanceof MissingApiKeyError ? 503 : 500;
      res.status(status).json({ error: formatGeminiError(error, 'Failed to generate stage section.') });
    }
  });

  app.post('/api/gemini/scrutinize-stage-all', async (req, res) => {
    try {
      const { docType, sections, sourceId, branchProposalSource, stageSource, context } = req.body as {
        docType: StageMemoDocType;
        sections?: StageScrutinySectionKey[];
        sourceId?: string;
        branchProposalSource?: IncomingBranchProposalSource | null;
        stageSource?: IncomingBranchProposalSource | null;
        context?: Record<string, unknown>;
      };

      if (!docType || !VALID_STAGE_DOCS.has(docType)) {
        res.status(400).json({ error: 'Invalid stage memo type.' });
        return;
      }

      const allowed = STAGE_SECTIONS_MAP[docType];
      const targetSections = (
        Array.isArray(sections) && sections.length > 0 ? sections : allowed
      ).filter((s): s is StageScrutinySectionKey => allowed.includes(s));

      if (targetSections.length === 0) {
        res.status(400).json({ error: 'No valid stage sections requested.' });
        return;
      }

      const source =
        resolveOptionalStageSource(sourceId, stageSource) ??
        resolveOptionalStageSource(sourceId, branchProposalSource);
      const aiClient = resolveAiClient();
      if (!aiClient.ok) {
        res.status(aiClient.status).json({ error: aiClient.error });
        return;
      }
      const ai = aiClient.ai;

      res.setHeader('Content-Type', 'text/event-stream; charset=utf-8');
      res.setHeader('Cache-Control', 'no-cache, no-transform');
      res.setHeader('Connection', 'keep-alive');
      res.setHeader('X-Accel-Buffering', 'no');
      res.flushHeaders?.();

      let clientClosed = false;
      res.on('close', () => {
        if (!res.writableEnded) clientClosed = true;
      });

      const writeEvent = (payload: Record<string, unknown>) => {
        if (clientClosed || res.writableEnded) return;
        res.write(`data: ${JSON.stringify(payload)}\n\n`);
      };

      const contextStr = buildCompactStageContext(context);
      const uncached: StageScrutinySectionKey[] = [];
      for (const sec of targetSections) {
        const cacheKey = `${source?.sourceId ?? 'ctx_only'}::STAGE::${sec}::${hashString(contextStr)}`;
        const cached = sectionResultCache.get(cacheKey);
        if (cached) {
          writeEvent({
            type: 'section',
            section: sec,
            result: cached.result,
            insufficient: cached.insufficient,
            insufficientReason: cached.insufficientReason,
          });
        } else {
          uncached.push(sec);
        }
      }

      if (uncached.length === 1 && !clientClosed) {
        const sec = uncached[0];
        try {
          const ev = await scrutinizeSingleStageSection(
            ai,
            sec,
            source,
            context,
            () => clientClosed,
            contextStr
          );
          writeEvent({
            type: 'section',
            section: sec,
            result: ev.result,
            insufficient: ev.insufficient,
            insufficientReason: ev.insufficientReason,
          });
        } catch (err) {
          if (!clientClosed) {
            writeEvent({
              type: 'error',
              section: sec,
              error: formatGeminiError(err, 'Failed to generate stage section.'),
            });
          }
        }
      } else if (uncached.length > 1 && !clientClosed) {
        const emittedStageKeys = new Set<string>();
        let stageBatchError: string | null = null;
        let hitHardQuota = false;

        const evaluateAndEmitStage = (sec: StageScrutinySectionKey, rawObj: Record<string, unknown>) => {
          const rawText = String(rawObj.text ?? '').trim();
          const cleanText = sanitizeAiMemoText(rawText);
          const modelReason = String(rawObj.insufficientReason ?? '').trim();
          const insufficient = Boolean(modelReason) || cleanText === '';
          const evaluated = {
            result: { text: cleanText, insufficientReason: modelReason },
            insufficient,
            insufficientReason:
              modelReason ||
              (!cleanText ? 'Required chained memo context or source data is unavailable.' : ''),
          };
          const cacheKey = `${source?.sourceId ?? 'ctx_only'}::STAGE::${sec}::${hashString(contextStr)}`;
          storeSectionCache(cacheKey, evaluated);
          writeEvent({
            type: 'section',
            section: sec,
            result: evaluated.result,
            insufficient: evaluated.insufficient,
            insufficientReason: evaluated.insufficientReason,
          });
        };

        try {
          const { prompt, schema } = buildConsolidatedStageBatchSpec(uncached, context);
          const textPayload = `${prompt}

${source ? `Uploaded Branch Proposal File: ${source.fileName}\nExtracted Text:\n${source.rawText}\n` : ''}
Chained Memo & Verified Document Context (IC Memo -> MEC Memo -> Board Memo -> Sanction + Stage Inputs):
${contextStr}`;

          const icAi = (context?.icMemoAiSections ?? {}) as Record<string, unknown>;
          const hasPopulatedIc = Boolean(icAi.justification || icAi.samdObservation);
          const shouldAttachImages =
            Boolean(source) &&
            (uncached.includes('sanctionCrossVerifiedSummary') ||
              !hasPopulatedIc ||
              (source?.rawText?.trim().length ?? 0) < 500);

          const parts: Array<
            | { inlineData: { mimeType: string; data: string } }
            | { text: string }
          > = shouldAttachImages && source
            ? [...source.imageParts, { text: textPayload }]
            : [{ text: textPayload }];

          const fullBuffer = await executeWithModelFallback(async (model) => {
            const stream = await ai.models.generateContentStream({
              model,
              contents: { parts },
              config: buildModelConfig(model, {
                systemInstruction: SCRUTINY_SYSTEM_INSTRUCTION,
                responseMimeType: 'application/json',
                responseSchema: schema,
              }),
            });

            let acc = '';
            for await (const chunk of stream) {
              if (clientClosed) break;
              const delta = chunk.text ?? '';
              if (!delta) continue;
              acc += delta;

              const completed = extractCompletedTopLevelObjects(acc, emittedStageKeys);
              for (const item of completed) {
                const sec = item.key as StageScrutinySectionKey;
                if (!uncached.includes(sec)) continue;
                evaluateAndEmitStage(sec, item.value);
              }
            }
            return acc;
          }, () => clientClosed);

          if (!clientClosed && fullBuffer.trim()) {
            const cleanedBuffer = stripJsonFences(fullBuffer);
            const salvaged = extractCompletedTopLevelObjects(cleanedBuffer, emittedStageKeys);
            for (const item of salvaged) {
              const sec = item.key as StageScrutinySectionKey;
              if (!uncached.includes(sec)) continue;
              evaluateAndEmitStage(sec, item.value);
            }
            try {
              const finalParsed = JSON.parse(cleanedBuffer) as Record<string, Record<string, unknown>>;
              for (const sec of uncached) {
                if (emittedStageKeys.has(sec)) continue;
                const rawObj = finalParsed[sec];
                if (rawObj && typeof rawObj === 'object') {
                  emittedStageKeys.add(sec);
                  evaluateAndEmitStage(sec, rawObj);
                }
              }
            } catch {
              // Remaining unemitted sections handled below
            }
          }
        } catch (err) {
          const rawMsg = err instanceof Error ? err.message : String(err);
          hitHardQuota = isQuotaOrRateError(rawMsg) && isDailyQuotaError(rawMsg);
          stageBatchError = formatGeminiError(err, 'Failed to generate stage section.');
        }

        const remainingStage = uncached.filter((sec) => !emittedStageKeys.has(sec));
        if (remainingStage.length > 0 && !clientClosed) {
          if (hitHardQuota && stageBatchError) {
            for (const sec of remainingStage) {
              writeEvent({ type: 'error', section: sec, error: stageBatchError });
            }
          } else {
            await runWithConcurrency(remainingStage, 2, async (sec) => {
              if (clientClosed) return;
              try {
                const ev = await scrutinizeSingleStageSection(
                  ai,
                  sec,
                  source,
                  context,
                  () => clientClosed,
                  contextStr
                );
                writeEvent({
                  type: 'section',
                  section: sec,
                  result: ev.result,
                  insufficient: ev.insufficient,
                  insufficientReason: ev.insufficientReason,
                });
              } catch (secErr) {
                if (!clientClosed) {
                  writeEvent({
                    type: 'error',
                    section: sec,
                    error: formatGeminiError(
                      secErr,
                      stageBatchError || 'Failed to generate stage section.'
                    ),
                  });
                }
              }
            });
          }
        }
      }

      if (!clientClosed) {
        writeEvent({ type: 'done', sourceId: source?.sourceId });
        res.end();
      }
    } catch (error) {
      if (!(error instanceof MissingApiKeyError)) {
        console.error('Error during stage batch scrutiny:', error);
      }
      const status = error instanceof MissingApiKeyError ? 503 : 500;
      const message = formatGeminiError(error, 'Stage batch generation failed.');
      if (!res.headersSent) {
        res.status(status).json({ error: message });
      } else if (!res.writableEnded) {
        res.write(`data: ${JSON.stringify({ type: 'fatal', error: message })}\n\n`);
        res.end();
      }
    }
  });

  // Legacy AI "Generate" sections endpoint (preserved for backwards compatibility)
  app.post('/api/gemini/generate-section', async (req, res) => {
    try {
      const { section, context } = req.body as {
        section: 'justification' | 'branchVisit' | 'samdObservation';
        context: Record<string, unknown>;
      };

      if (!section || !context) {
        res.status(400).json({ error: 'Missing section type or memo context.' });
        return;
      }

      const aiClient = resolveAiClient();
      if (!aiClient.ok) {
        res.status(aiClient.status).json({ error: aiClient.error });
        return;
      }
      const ai = aiClient.ai;

      const systemInstruction =
        'You are a Senior Credit Analyst at the Special Asset Management Division (SAMD), Head Office, Sammilito Islami Bank PLC (Union Bank PLC Unit), Dhaka. Write formal, concise, objective banking memorandum prose strictly grounded in the provided document data and figures in Million BDT. Do not invent figures or facts not present in the context.';

      let userPrompt = '';

      if (section === 'justification') {
        userPrompt = `Scrutinize the IC Memo data (placeholders, settlement/waiver figures, Cost of Fund, security, classification, source of repayment) and all relevant details of the Branch Proposal below, and write exactly 3 key and relevant justification points for placeholder #51# under "Justification of The Proposed Facility".
Format as exactly 3 numbered justification points separated by newlines ("1. ...", "2. ...", "3. ..."), covering: (1) client's current business/financial position and genuine distress without willful default, (2) recovery & settlement math (down payment / principal / cost of fund coverage & protection of bank interest), and (3) immediate reduction of classified/NPL exposure and avoidance of prolonged litigation.
Keep each point crisp, factual, and substantive (1-2 sentences each). Do NOT use markdown bolding or asterisks.

Context:
${JSON.stringify(context, null, 2)}`;
      } else if (section === 'branchVisit') {
        userPrompt = `Scrutinize all relevant sections of the Branch Proposal (including Section 12 Branch Visit Report, Section 10 Client Particulars, office/factory addresses, last visit date, stock position, fixed assets, business net worth, Section 7 security details, Section 11 financials, and Section 13/15/16 branch recommendations) and the IC Memo client particulars below, and write a comprehensive Section 12 "Branch Visit Report" for the Head Office IC Memo.
CRITICAL REQUIREMENT: Write from the perspective of the branch while strictly treating the branch in the THIRD PERSON (e.g., "The officials of [Branch Name] Branch...", "Branch management conducted an inspection...", "The Branch observed that...", never use "we", "us", "our", or "our branch").
Include the visiting officials' names/designations if available, visit date, client/concern name, premises/factory/collateral inspected, physical stock/asset & operational business findings, client interaction, and branch verification outcome accurately. Write 1 to 2 comprehensive, formal banking paragraphs. Do NOT use markdown bolding or asterisks.

Context:
${JSON.stringify(context, null, 2)}`;
      } else if (section === 'samdObservation') {
        userPrompt = `Thoroughly scrutinize the IC Memo figures and sections (Principal, Charges, Charged Profit, Compensation, Uncharged Compensation, NET Outstanding, Cost of Fund Surplus/Deficit, Maximum Waiver Considering COF #44#, Down Payment, Proposed Waiver, Settlement/Exit Amount, Settlement Period, CIB & Classification status, Legal Action status, Audit/Regulatory status, Security MV/FSV, Source of Repayment, Branch Visit Report) and the relevant details of the Branch Proposal below, and write Section 15 "Observation and Recommendation of SAMD".
Write exactly five (5) key and substantive justification points separated by newlines ("1. ...", "2. ...", "3. ...", "4. ...", "5. ...") reflecting SAMD's observations and recommendations, addressing:
1. Verification of total outstanding liability breakup, classification/overdue/CIB status, and confirmation of absence of fraud, forgery, or willful default.
2. Assessment of primary & collateral security valuation (MV/FSV coverage), legal action status, and audit/regulatory compliance position.
3. Cost of Fund & waiver adequacy analysis (confirming down payment, recovery covering principal/charges/Cost of Fund, maximum permissible waiver considering COF, and conformity with BRPD Circular Letter No. 23 dated 29.06.2026).
4. Evaluation of client's genuine business/financial condition, branch visit & branch recommendation findings, and viability of the proposed source of repayment within the settlement deadline.
5. Formal SAMD recommendation to place the Exit with Waiver proposal before the Investment Committee for exit within the stipulated settlement period subject to compliance with special terms and conditions.
Each point should be 1-2 substantive, formal banking sentences. Do NOT use markdown bolding or asterisks.

Context:
${JSON.stringify(context, null, 2)}`;
      } else {
        res.status(400).json({ error: 'Unknown section identifier.' });
        return;
      }

      const response = await executeWithModelFallback((model) =>
        ai.models.generateContent({
          model,
          contents: userPrompt,
          config: buildModelConfig(model, { systemInstruction }),
        })
      );

      const text = (response.text ?? '').trim();
      res.json({ text });
    } catch (error) {
      if (!(error instanceof MissingApiKeyError)) {
        console.error('Error generating IC Memo section:', error);
      }
      const status = error instanceof MissingApiKeyError ? 503 : 500;
      const message = formatGeminiError(error, 'Failed to generate section.');
      res.status(status).json({ error: message });
    }
  });

  // Guarantee /api/* requests never fall through to Vite/SPA HTML fallback or Express HTML error pages
  app.use('/api', (_req, res) => {
    res.status(404).json({ error: 'API endpoint not found.' });
  });

  app.use(
    (
      err: unknown,
      req: express.Request,
      res: express.Response,
      next: express.NextFunction
    ) => {
      if (!req.path.startsWith('/api')) {
        next(err);
        return;
      }
      if (res.headersSent) {
        next(err);
        return;
      }
      const errObj = err as { status?: number; statusCode?: number; type?: string; message?: string };
      const status =
        typeof errObj?.status === 'number'
          ? errObj.status
          : typeof errObj?.statusCode === 'number'
            ? errObj.statusCode
            : 500;
      const message =
        errObj?.type === 'entity.too.large' || status === 413
          ? 'Uploaded document payload exceeds the server request limit. Please reduce file size or page count and retry.'
          : errObj?.type === 'entity.parse.failed'
            ? 'Invalid JSON request payload.'
            : errObj?.message || 'Unexpected API server error.';
      res.status(status).json({ error: message });
    }
  );

  const httpServer = http.createServer(app);

  if (process.env.NODE_ENV !== 'production') {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: {
        middlewareMode: true,
        allowedHosts: true,
        ws: {
          server: httpServer,
        },
      },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.resolve(__dirname, 'dist');
    app.use(express.static(distPath));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  httpServer.listen(PORT, '0.0.0.0', () => {
    console.log(`Memo Engine by Tuhin running on http://0.0.0.0:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error('Failed to start server:', err);
  process.exit(1);
});
