import type {
  BranchProposalData,
  Section3Row,
  Section4Row,
  Section5Row,
  Section6Row,
  Section7PrimaryRow,
  Section7ValuationRow,
  Section9Row,
  Section10BankDirectorRow,
  Section10DirectorRow,
  Section10Particulars,
  Section11Row,
} from '../types/memo';
import { cleanText } from '../utils/numberFormat';
import type { PdfLine, PdfTextItem } from './pdfUtils';

function emptyParticulars(): Section10Particulars {
  return {
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
}

export function createEmptyBranchProposal(): BranchProposalData {
  return {
    referenceNo: '',
    proposalDate: '',
    branchName: '',
    accountNo: '',
    subjectText: '',
    limitOriginal: '',
    limitRSD: '',
    securityMV: '',
    securityFSV: '',
    justificationText: '',
    previousBoardDecisions: '',
    natureOfBusiness: '',
    sec2Limit: '',
    sec2Validity: '',
    sec2Overdue: '',
    sec2Recovery: '',
    sec2Remarks: '',
    section3AsOnDate: '',
    section3Rows: [],
    section4CibDate: '',
    section4Rows: [],
    section5StatusNote: '',
    section5Rows: [],
    section6StatusNote: '',
    section6Rows: [],
    section7PrimaryNote: '',
    section7PrimaryRows: [],
    section7PrimaryTotal: '',
    section7CollateralNote: '',
    section7CollateralSchedules: [],
    section7ValuationRows: [],
    section7CoveragePct: '',
    section7Others: '',
    section8SourceOfRepayment: '',
    section9Rows: [],
    section10Particulars: emptyParticulars(),
    section10Directors: [],
    section10BankDirectors: [],
    section11FinancialsNote: '',
    section11Rows: [],
    section12BranchVisitReport: '',
    section15BranchBicRec: '',
    section16BranchRec: '',
  };
}

/** Normalize subtle OCR homoglyphs inside candidate numeric strings (e.g. "0.s0" -> "0.50") */
function normalizeOcrNumericToken(raw: string): string | null {
  let s = cleanText(raw)
    .replace(/\bmillion\b/gi, '')
    .replace(/tk\.?/gi, '')
    .replace(/bdt/gi, '')
    .trim();
  if (!s) return null;

  // Handle "0.s 0" or "0.,50" or "0.s0"
  s = s.replace(/\s+/g, '');
  s = s.replace(/\.,|,\./g, '.');
  // Only convert OCR confusable letters when adjacent to digits or dot
  if (/[0-9]/.test(s)) {
    s = s
      .replace(/[sS]/g, '5')
      .replace(/[Oo]/g, '0')
      .replace(/[Il|]/g, '1')
      .replace(/tJ/g, '8');
  }
  const match = s.match(/^-?\d+(?:,\d+)*(?:\.\d+)?$/);
  if (!match) return null;
  const num = Number(s.replace(/,/g, ''));
  if (!Number.isFinite(num)) return null;
  return s;
}

/** Normalize OCR spacing/homoglyphs in date tokens like "01.0 9 .2 0 25" or "t0.11 .2026" */
function extractDateToken(raw: string): string {
  let collapsed = raw.replace(/\s*([./-])\s*/g, '$1');
  while (/\d\s+\d/.test(collapsed)) {
    collapsed = collapsed.replace(/(\d)\s+(\d)/g, '$1$2');
  }
  collapsed = collapsed
    .replace(/\b[tI](\d[./-]\d{1,2}[./-]\d{2,4})\b/g, '1$1')
    .replace(/\.I\s*1\./g, '.11.');
  const m =
    collapsed.match(/(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})/) ||
    collapsed.match(/(\d{1,2}-[A-Za-z]{3}-\d{2,4})/) ||
    collapsed.match(/\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}\b/i);
  return m ? cleanText(m[1] ?? m[0]) : '';
}

function normalizeCommonOcrWords(text: string): string {
  return cleanText(text)
    .replace(/\bl0\/-/g, '10/-')
    .replace(/\b21s675\b/gi, '215675')
    .replace(/\b03-5\s*I\s*I\s*10034278\b/gi, '03-51110034278')
    .replace(/\b035111003,1278\b/g, '0351110034278')
    .replace(/\b03-511100i,1278\b/gi, '03-51110034278')
    .replace(/\bBI\{PD/g, 'BRPD')
    .replace(/\bNovernber\b/gi, 'November')
    .replace(/\bSarne\b/gi, 'Same')
    .replace(/\b0\.,50\b/g, '0.50')
    .replace(/\bN\/?A\b/gi, 'N/A');
}

function slicePdfLinesBetween(
  lines: PdfLine[],
  startRe: RegExp,
  endRe: RegExp | null,
  fromIndex = 0
): { slice: PdfLine[]; startIdx: number; endIdx: number } {
  let startIdx = -1;
  for (let i = fromIndex; i < lines.length; i++) {
    if (startRe.test(lines[i].text)) {
      startIdx = i;
      break;
    }
  }
  if (startIdx === -1) return { slice: [], startIdx: -1, endIdx: -1 };

  let endIdx = lines.length;
  if (endRe) {
    for (let i = startIdx + 1; i < lines.length; i++) {
      if (endRe.test(lines[i].text)) {
        endIdx = i;
        break;
      }
    }
  }
  return { slice: lines.slice(startIdx, endIdx), startIdx, endIdx };
}

function itemsInXBand(items: PdfTextItem[], minX: number, maxX: number): string {
  return cleanText(
    items
      .filter((it) => {
        const cx = it.x + Math.min(it.width, 40) / 2;
        return cx >= minX && cx < maxX;
      })
      .map((it) => it.text.trim())
      .filter(Boolean)
      .join(' ')
  );
}

function isPageHeaderFooterLine(text: string): boolean {
  return (
    /sammilito\s+islami\s+bank/i.test(text) ||
    /state-?o[wu]?ned.*shariah/i.test(text) ||
    /union\s+bank\s+plc\s+unit/i.test(text) ||
    /^exit\s+with\s+waiver$/i.test(text.trim())
  );
}

/**
 * Rule-based non-AI extractor for Branch Proposal PDFs (or OCR-converted PdfLine arrays).
 * Reads directly from PDF text coordinates in milliseconds without any AI API call.
 */
export function parseBranchProposalFromPdf(
  lines: PdfLine[],
  rawText: string
): BranchProposalData {
  const bp = createEmptyBranchProposal();
  if (!lines.length && !rawText.trim()) return bp;

  const p1Lines = lines.filter((l) => l.page === 1);
  const topMetaLines = p1Lines.slice(0, 22);

  // 1. Reference No (#64#), Date (#65#), Branch, A/C No, Subject
  for (const l of topMetaLines) {
    const t = cleanText(l.text);
    if (!bp.referenceNo) {
      const refM = t.match(/Ref(?:\.?\s*No\.?)?\s*:?\s*([A-Z]{2,}[A-Z0-9.-]*(?:\/[A-Z0-9.-]+){2,})/i);
      if (refM && !/[,\\'+]/.test(refM[1])) {
        bp.referenceNo = refM[1].trim();
      } else {
        const bareRef = t.match(/\b([A-Z]{2,6}\/[A-Z0-9.-]+\/[A-Z0-9.-]+\/\d{4}\/[A-Z0-9-]+)\b/);
        if (bareRef && !/[,\\'+]/.test(bareRef[1])) {
          bp.referenceNo = bareRef[1].trim();
        }
      }
    }

    if (!bp.proposalDate && /\bDate\s*:/i.test(t) && !/circular/i.test(t)) {
      const d = extractDateToken(t.replace(/^.*Date\s*:/i, ''));
      if (d && /\d{1,2}[./-]\d{1,2}[./-]\d{4}/.test(d)) {
        bp.proposalDate = d;
      }
    }

    if (!bp.branchName) {
      const brM = t.match(/\bBranch\s*:\s*(.+)$/i);
      if (brM) {
        bp.branchName = cleanText(brM[1]).replace(/_+/g, '');
      } else {
        const unitBr = t.match(/Unit[.,]?\s*([A-Za-z]+(?:\s+[A-Za-z]+)?)\s+Branch/i);
        if (unitBr) bp.branchName = `${cleanText(unitBr[1])} Branch`;
      }
    }

    if (!bp.accountNo) {
      const acM = t.match(/\bA\/C(?:\s*No\.?)?\s*:?\s*([0-9A-Z-]{8,20})/i);
      if (acM) {
        bp.accountNo = cleanText(acM[1]).replace(/O/g, '0').replace(/J/g, '3');
      }
    }
  }

  // Subject block
  const subjStart = p1Lines.findIndex((l) => /Subject\s*:/i.test(l.text));
  if (subjStart !== -1) {
    const subjLines: string[] = [];
    for (let i = subjStart; i < Math.min(p1Lines.length, subjStart + 8); i++) {
      const t = p1Lines[i].text;
      if (/^i\.|^the\s+client\s+will|figures\s+in\s+this|^\bmode\b/i.test(t)) break;
      subjLines.push(t.replace(/^Subject\s*:\s*/i, ''));
    }
    bp.subjectText = cleanText(subjLines.join(' '));
  }

  // 2. Main Settlement Table (#6# limitOriginal, limitRSD, #33# MV, #34# FSV)
  const mainHdrIdx = p1Lines.findIndex(
    (l) =>
      (/Limit|O\/S\s+as\s+on/i.test(l.text) && /Settlement|Waiver/i.test(l.text)) ||
      /0?riginal\s+RS[DO)]/i.test(l.text)
  );
  const mainEndIdx = p1Lines.findIndex(
    (l, idx) => idx > Math.max(0, mainHdrIdx) && (/Note\s*:|Key\s+Points/i.test(l.text))
  );
  if (mainHdrIdx !== -1) {
    const end = mainEndIdx > mainHdrIdx ? mainEndIdx : Math.min(p1Lines.length, mainHdrIdx + 18);
    const tableLines = p1Lines.slice(mainHdrIdx, end);

    let origMinX = 110;
    let origMaxX = 165;
    let rsdMinX = 165;
    let rsdMaxX = 200;
    let remMinX = 485;

    for (const tl of tableLines.slice(0, 4)) {
      const origItem = tl.items.find((it) => /0?riginal/i.test(it.text));
      const rsdItem = tl.items.find((it) => /RS[DOI)]/i.test(it.text));
      const remItem = tl.items.find((it) => /Remar[hk]/i.test(it.text));
      if (origItem) {
        origMinX = origItem.x - 10;
        origMaxX = rsdItem ? rsdItem.x - 4 : origItem.x + 48;
      }
      if (rsdItem) {
        rsdMinX = rsdItem.x - 4;
        rsdMaxX = rsdItem.x + 32;
      }
      if (remItem) remMinX = remItem.x - 12;
    }

    for (const tl of tableLines.slice(1)) {
      const origRaw = itemsInXBand(tl.items, origMinX, origMaxX);
      if (origRaw && !bp.limitOriginal && !/0?riginal|limit|^01$/i.test(origRaw)) {
        const num = normalizeOcrNumericToken(origRaw);
        if (num) bp.limitOriginal = num;
      }
      const rsdRaw = itemsInXBand(tl.items, rsdMinX, rsdMaxX);
      if (rsdRaw && !bp.limitRSD && !/RS[DOI)]/i.test(rsdRaw)) {
        bp.limitRSD = rsdRaw;
      }

      const remText = itemsInXBand(tl.items, remMinX, 900);
      if (remText) {
        const mvM = remText.match(/MV\s*:\s*([A-Za-z0-9.,/-]+)/i);
        if (mvM && !bp.securityMV) bp.securityMV = cleanText(mvM[1]);
        const fsvM = remText.match(/FSV\s*:\s*([A-Za-z0-9.,/-]+)/i);
        if (fsvM && !bp.securityFSV) bp.securityFSV = cleanText(fsvM[1]);
      }
    }
  }

  // Fallback check across table lines for MV:/FSV:
  if (!bp.securityMV) {
    const mvLine = p1Lines.find((l) => /\bMV\s*:/i.test(l.text));
    if (mvLine) {
      const m = mvLine.text.match(/\bMV\s*:\s*(N\/A|[\d.,]+)/i);
      if (m) bp.securityMV = cleanText(m[1]);
    }
  }
  if (!bp.securityFSV) {
    const fsvLine = p1Lines.find((l) => /\bFSV\s*:/i.test(l.text));
    if (fsvLine) {
      const m = fsvLine.text.match(/\bFSV\s*:\s*(N\/A|[\d.,]+)/i);
      if (m) bp.securityFSV = cleanText(m[1]);
    }
  }

  // 3. Justification & Previous Board Decisions on Page 1
  const justSec = slicePdfLinesBetween(
    p1Lines,
    /J\s*ust[a-z']*\s+o[frl']*\s+(?:The\s+)?.*Facilit/i,
    /Previous\s+Board\s+Decision|^\s*1\s*\./i
  );
  if (justSec.slice.length > 0) {
    const body = justSec.slice
      .slice(1)
      .map((l) => l.text)
      .filter((t) => !/^['"]?[FT]k\.?$/i.test(t.trim()));
    bp.justificationText = cleanText(body.join(' '));
  }

  const prevBoardLine = lines.find((l) => /Previous\s+Board\s+Decision/i.test(l.text));
  if (prevBoardLine) {
    const m = prevBoardLine.text.match(/Previous\s+Board\s+Decision[^:]*:\s*(.+)/i);
    if (m) bp.previousBoardDecisions = cleanText(m[1]);
  }

  // 4. Section 1: Nature of Business (#54#)
  const sec1Explicit = lines.find((l) =>
    /(?:^|\b)1\s*\.\s*Nature\s+of\s+Bus[a-z']*s?\s*:/i.test(l.text)
  );
  if (sec1Explicit) {
    const m = sec1Explicit.text.match(/Nature\s+of\s+Bus[a-z']*s?\s*:\s*(.+)/i);
    if (m) bp.natureOfBusiness = cleanText(m[1]);
  } else {
    const sec2Idx = lines.findIndex((l) => /^\s*2\s*\b.*(?:Our\s+Bank|Liability)/i.test(l.text));
    if (sec2Idx > 0) {
      for (let i = Math.max(0, sec2Idx - 4); i < sec2Idx; i++) {
        const txt = lines[i].text;
        if (/ricu\s*[lI]tu\s*re/i.test(txt)) {
          bp.natureOfBusiness = 'Agriculture';
          break;
        }
        const m = txt.match(/^\s*[1Il]\s*\.\s*([A-Za-z][A-Za-z &/-]{2,})$/);
        if (m && !isPageHeaderFooterLine(m[1])) {
          bp.natureOfBusiness = cleanText(m[1]);
          break;
        }
      }
    }
  }

  // 5. Section 2: Liability Position (#56#, #58#, sec2Overdue, sec2Recovery, sec2Remarks)
  const sec2 = slicePdfLinesBetween(
    lines,
    /^\s*2\s*\b.*(?:Our\s+Bank|Liability)/i,
    /^\s*3\s*\.\s*/i
  );
  if (sec2.slice.length > 0) {
    const hdrIdx = sec2.slice.findIndex((l) => /Limit/i.test(l.text) && /Validit/i.test(l.text));
    if (hdrIdx !== -1) {
      const hdr = sec2.slice[hdrIdx];
      const limitIt = hdr.items.find((it) => /Limit/i.test(it.text));
      const osIt = hdr.items.find((it) => /Outstand/i.test(it.text));
      const valIt = hdr.items.find((it) => /Validit/i.test(it.text));
      const ovdIt = hdr.items.find((it) => /Overdue/i.test(it.text));
      const recIt = hdr.items.find((it) => /Recover|[Il]lccover/i.test(it.text));
      const remIt = hdr.items.find((it) => /Remar|[Il]lc[nu]a/i.test(it.text));

      const limX = limitIt ? limitIt.x - 8 : 225;
      const osX = osIt ? osIt.x - 8 : 270;
      const valX = valIt ? valIt.x - 8 : 332;
      const ovdX = ovdIt ? ovdIt.x - 8 : 385;
      const recX = recIt ? recIt.x - 8 : 432;
      const remX = remIt ? remIt.x - 8 : 500;

      const dataRows = sec2.slice
        .slice(hdrIdx + 1)
        .filter((l) => !/^(?:Funded|Non\s*-?\s*Funded|Total|Investment|Concern)\b/i.test(l.text));

      const limParts: string[] = [];
      const valParts: string[] = [];
      const ovdParts: string[] = [];
      const recParts: string[] = [];
      const remParts: string[] = [];

      for (const dl of dataRows) {
        if (/\bTotal\b/i.test(dl.text)) break;
        const lTxt = itemsInXBand(dl.items, limX, osX);
        const vTxt = itemsInXBand(dl.items, valX, ovdX);
        const oTxt = itemsInXBand(dl.items, ovdX, recX);
        const rTxt = itemsInXBand(dl.items, recX, remX);
        const mTxt = itemsInXBand(dl.items, remX, 900);

        if (lTxt && !/Investment/i.test(lTxt)) limParts.push(lTxt);
        if (vTxt && !/date/i.test(vTxt)) valParts.push(vTxt);
        if (oTxt) ovdParts.push(oTxt);
        if (rTxt && !/deals|till|Mode|1'tse/i.test(rTxt)) recParts.push(rTxt);
        if (mTxt && !/Ren|as/i.test(mTxt)) remParts.push(mTxt);
      }

      if (limParts.length > 0) {
        const joinedLim = cleanText(limParts.join(' '));
        const num = normalizeOcrNumericToken(joinedLim);
        bp.sec2Limit = num || joinedLim;
      }
      if (valParts.length > 0) {
        const joinedVal = cleanText(valParts.join(' '));
        const d = extractDateToken(joinedVal);
        const hasExp = /ex/i.test(joinedVal);
        bp.sec2Validity = d ? (hasExp ? `${d} (Exp.)` : d) : joinedVal;
      }
      if (ovdParts.length > 0) bp.sec2Overdue = cleanText(ovdParts.join(' '));
      if (recParts.length > 0) {
        const recClean = cleanText(recParts.join(' ')).replace(/\\1'tse/g, '');
        const recNum = normalizeOcrNumericToken(recClean);
        bp.sec2Recovery = recNum || recClean;
      }
      if (remParts.length > 0) bp.sec2Remarks = cleanText(remParts.join(' '));
    }
  }

  // 6. Section 3: Liability with Other Units of SIB / Other Banks
  const sec3 = slicePdfLinesBetween(
    lines,
    /^\s*3\s*\.\s*/i,
    /^\s*4\s*\.\s*CIB/i
  );
  if (sec3.slice.length > 0) {
    const asOn = extractDateToken(sec3.slice[0].text);
    if (asOn) bp.section3AsOnDate = asOn;

    const dataStart = sec3.slice.findIndex((l) => /^\s*a\.?\b/i.test(l.text));
    if (dataStart !== -1) {
      const rows: Section3Row[] = [];
      for (const rl of sec3.slice.slice(dataStart)) {
        if (/sub[-\s]*total/i.test(rl.text)) {
          rows.push({
            sl: '',
            nameOfConcern: 'Sub-Total',
            bankOrNbfi: '',
            natureOfFacility: '',
            limit: itemsInXBand(rl.items, 310, 360),
            outstanding: itemsInXBand(rl.items, 360, 445),
            overdueEol: itemsInXBand(rl.items, 445, 505),
            expiryDate: '',
            clStatus: '',
          });
          continue;
        }
        if (/grand\s*total/i.test(rl.text)) {
          rows.push({
            sl: '',
            nameOfConcern: 'Grand Total',
            bankOrNbfi: '',
            natureOfFacility: '',
            limit: itemsInXBand(rl.items, 310, 360),
            outstanding: itemsInXBand(rl.items, 360, 445),
            overdueEol: itemsInXBand(rl.items, 445, 505),
            expiryDate: '',
            clStatus: '',
          });
          continue;
        }
        const slM = rl.text.match(/^([a-z]\.?|\d+\.?)\b/i);
        if (slM) {
          rows.push({
            sl: slM[1],
            nameOfConcern: itemsInXBand(rl.items, 95, 172),
            bankOrNbfi: itemsInXBand(rl.items, 172, 258),
            natureOfFacility: itemsInXBand(rl.items, 258, 312),
            limit: itemsInXBand(rl.items, 312, 358),
            outstanding: itemsInXBand(rl.items, 358, 446),
            overdueEol: itemsInXBand(rl.items, 446, 505),
            expiryDate: itemsInXBand(rl.items, 505, 550),
            clStatus: itemsInXBand(rl.items, 550, 900),
          });
        }
      }
      bp.section3Rows = rows;
    }
  }

  // 7. Section 4: CIB Status
  const sec4 = slicePdfLinesBetween(
    lines,
    /^\s*4\s*\.\s*CIB/i,
    /^\s*5\s*(?:\.|\s+Legal|\b)/i
  );
  if (sec4.slice.length > 0) {
    const cibD = extractDateToken(sec4.slice[0].text);
    if (cibD) bp.section4CibDate = cibD;

    const statusMatchers: Array<{ label: string; re: RegExp }> = [
      { label: 'STD', re: /^(?:STD|SI'D|S'FD)$/i },
      { label: 'SMA', re: /^SMA$/i },
      { label: 'SS', re: /^(?:SS|55)$/i },
      { label: 'DF', re: /^(?:DF|I\)F|DF,)$/i },
      { label: 'BL', re: /^(?:BL|I\]L|I\]I,|RL|BI,)$/i },
      { label: 'Willful Default (WD)', re: /Willful\s*Defa/i },
      { label: 'Write-Off', re: /write[- ]?o[flf]{2}/i },
      { label: 'Non-Funded', re: /Non[- ]?[Ff]r?unded/i },
      { label: 'Stay Order', re: /Sta[vy]?\s*Order/i },
      { label: 'TOTAL', re: /^TOTAL\b/i },
    ];

    const sec4Rows: Section4Row[] = statusMatchers.map(({ label, re }) => {
      const matchedLine = sec4.slice.find((l) => {
        const left = itemsInXBand(l.items, 60, 185);
        return re.test(left) || re.test(l.text.split(/\s+/)[0] ?? '');
      });
      if (!matchedLine) {
        return {
          status: label,
          unitBorrower: '',
          unitGuarantor: '',
          otherBorrower: '',
          otherGuarantor: '',
          remarks: '',
        };
      }
      return {
        status: label,
        unitBorrower: normalizeCommonOcrWords(itemsInXBand(matchedLine.items, 185, 250))
          .replace(/^\)\s*/, '')
          .replace(/^222\s+/, ''),
        unitGuarantor: normalizeCommonOcrWords(itemsInXBand(matchedLine.items, 250, 312)),
        otherBorrower: normalizeCommonOcrWords(itemsInXBand(matchedLine.items, 312, 375))
          .replace(/\bNit\b/g, 'Nil'),
        otherGuarantor: normalizeCommonOcrWords(itemsInXBand(matchedLine.items, 375, 455)),
        remarks: normalizeCommonOcrWords(itemsInXBand(matchedLine.items, 455, 900)),
      };
    });
    bp.section4Rows = sec4Rows;
  }

  // 8. Section 5: Legal Action Status
  const sec5 = slicePdfLinesBetween(
    lines,
    /^\s*5\s*(?:\.|\s+Legal|$)/i,
    /^\s*6\s*\.\s*Audit/i
  );
  if (sec5.slice.length > 0) {
    const hdr = sec5.slice[0].text;
    const noteM = hdr.match(/(?:N\/A|Nil)/i);
    if (noteM) bp.section5StatusNote = noteM[0];
    const dataRows: Section5Row[] = [];
    for (const l of sec5.slice.slice(1)) {
      if (/Type\s*of\s*Case|Case\s*No|Filling/i.test(l.text)) continue;
      const sl = itemsInXBand(l.items, 60, 100);
      const typeOfCase = itemsInXBand(l.items, 100, 178);
      const caseNo = itemsInXBand(l.items, 178, 268);
      const filingDate = itemsInXBand(l.items, 268, 312);
      const amount = itemsInXBand(l.items, 312, 365);
      const previousDate = itemsInXBand(l.items, 365, 432);
      const nextDate = itemsInXBand(l.items, 432, 502);
      const presentStatus = itemsInXBand(l.items, 502, 900);
      if (typeOfCase || caseNo || amount) {
        dataRows.push({
          sl,
          typeOfCase,
          caseNo,
          filingDate,
          amount,
          previousDate,
          nextDate,
          presentStatus,
        });
      }
    }
    bp.section5Rows = dataRows;
  }

  // 9. Section 6: Audit, Compliance and Regulatory Authority
  const sec6 = slicePdfLinesBetween(
    lines,
    /^\s*6\s*\.\s*Audit/i,
    /^\s*7\s*\.\s*Securit/i
  );
  if (sec6.slice.length > 0) {
    const noteM = sec6.slice[0].text.match(/\b(N\/A|Nil)\b/i);
    if (noteM) bp.section6StatusNote = noteM[1];

    const sec6Defs: Array<{ sl: string; authority: string; match: RegExp }> = [
      { sl: 'a.', authority: 'IC & CD', match: /IC\s*&?\s*CD/i },
      { sl: 'b.', authority: 'DBI, Bangladesh Bank', match: /DBI|I\)t\]I|Bangladesh\s*Bank/i },
      { sl: 'c.', authority: 'DOS, Bangladesh Bank', match: /\bDOS\b/i },
      { sl: 'd.', authority: 'BFIU, Bangladesh Bank', match: /BFIU|RI\]IIJ/i },
      { sl: 'e', authority: 'BRD, Bangladesh Bank', match: /^\s*e\b/i },
      { sl: 'f', authority: 'ACC(Anti-Corruption Commission)', match: /ACC|Anti-Cor/i },
    ];

    const rows6: Section6Row[] = [];
    for (const def of sec6Defs) {
      const idx = sec6.slice.findIndex((l) => def.match.test(l.text));
      if (idx !== -1) {
        const rowLine = sec6.slice[idx];
        const insp = cleanText(
          itemsInXBand(rowLine.items, 200, 280).replace(/cloney?/gi, 'done')
        );
        const objStat = cleanText(itemsInXBand(rowLine.items, 280, 355).replace(/\bI\b/g, ''));
        const objs = cleanText(itemsInXBand(rowLine.items, 355, 490));
        const reply = cleanText(itemsInXBand(rowLine.items, 490, 900));
        rows6.push({
          sl: def.sl,
          authority: def.authority,
          lastInspectionDate: insp,
          objectionStatus: objStat,
          objections: objs,
          branchReply: reply,
        });
      }
    }
    if (rows6.some((r) => r.lastInspectionDate || r.objectionStatus || r.objections || r.branchReply)) {
      bp.section6Rows = rows6;
    }
  }

  // 10. Section 7: Security
  const sec7 = slicePdfLinesBetween(
    lines,
    /^\s*7\s*\.\s*Securit/i,
    /^\s*(?:8|9)\b/i
  );
  if (sec7.slice.length > 0) {
    const primNoteLine = sec7.slice.find((l) => /\ba\.\s*Pri/i.test(l.text));
    if (primNoteLine) {
      const m = primNoteLine.text.match(/\b(N\/A|Nil)\b/i);
      if (m) bp.section7PrimaryNote = m[1];
    }
    const collNoteLine = sec7.slice.find((l) => /\bb\.\s*[,(]*[Co]ll/i.test(l.text));
    if (collNoteLine) {
      const m = collNoteLine.text.match(/\b(N\/A|Nil)\b/i);
      if (m) bp.section7CollateralNote = m[1];
    }

    const covLine = sec7.slice.find((l) => /Coverage/i.test(l.text));
    if (covLine) {
      const m = covLine.text.match(/(\d+(?:\.\d+)?\s*%)/);
      if (m) bp.section7CoveragePct = m[1].replace(/\s+/g, '');
    }
  }

  // 11. Section 8: Source of Repayment
  const sec8 = slicePdfLinesBetween(
    lines,
    /^\s*8\s*\.\s*Source\s+of\s+Repayment/i,
    /^\s*9\b/i
  );
  if (sec8.slice.length > 0) {
    const text8 = cleanText(
      sec8.slice
        .map((l) => l.text)
        .join(' ')
        .replace(/^8\s*\.\s*Source\s+of\s+Repayment\s*:?\s*/i, '')
    );
    if (text8) bp.section8SourceOfRepayment = text8;
  }

  // 12. Section 9: Sequence of Head Office Sanction(S)
  const sec9 = slicePdfLinesBetween(
    lines,
    /^\s*9\s*(?:\.|\s+Sequence|\s+of\b)/i,
    /^\s*10\s*\.\s*/i
  );
  if (sec9.slice.length > 0) {
    const hdrIdx = sec9.slice.findIndex((l) => /Sanction|Reference/i.test(l.text));
    const dataLines = sec9.slice.slice(hdrIdx !== -1 ? hdrIdx + 1 : 1);
    const sec9Rows: Section9Row[] = [];

    for (const dl of dataLines) {
      const slRaw = itemsInXBand(dl.items, 55, 84);
      const slClean = /^[a-g]\.?$/i.test(slRaw)
        ? slRaw.endsWith('.') ? slRaw.toLowerCase() : `${slRaw.toLowerCase()}.`
        : /^il\.?$/i.test(slRaw)
          ? 'a.'
          : /^[Cc]\.?$/i.test(slRaw)
            ? 'c.'
            : '';
      if (!slClean) continue;

      const refRaw = itemsInXBand(dl.items, 84, 275);
      const dateRaw = extractDateToken(itemsInXBand(dl.items, 275, 338));
      const limRaw = itemsInXBand(dl.items, 338, 435);
      const remRaw = itemsInXBand(dl.items, 435, 900);

      let refClean = cleanText(refRaw);
      if (/SANC/i.test(refClean) && /\d{4}/.test(refClean) && /^u\s/i.test(refClean)) {
        refClean = refClean
          .replace(/^u\s+llt\.'l\s+Iolil\{M\s+I\),'?/i, 'UBL/HO/IRM/')
          .replace(/SANC,(\d{4})in?(\d+)/i, 'SANC/$1/$2');
      }
      const limNum = normalizeOcrNumericToken(limRaw);
      const limitStr = limNum ? `${limNum}${/million/i.test(limRaw) ? ' Million' : ''}` : cleanText(limRaw);

      sec9Rows.push({
        sl: slClean,
        referenceNo: refClean,
        date: dateRaw,
        limit: limitStr,
        remarks: cleanText(remRaw),
      });
    }
    if (sec9Rows.some((r) => r.referenceNo || r.date || r.limit || r.remarks)) {
      bp.section9Rows = sec9Rows;
    }
  }

  // 13. Section 10: Brief Particulars of The Client
  const sec10 = slicePdfLinesBetween(
    lines,
    /^\s*10\s*\.\s*/i,
    /^\s*(?:11|I\s*l)\s*\.\s*F/i
  );
  if (sec10.slice.length > 0) {
    const dirIdx = sec10.slice.findIndex((l) => /Particular\s+of\s+the\s+(?:Proprietor|Partner|Director)/i.test(l.text));
    const partLines = dirIdx !== -1 ? sec10.slice.slice(1, dirIdx) : sec10.slice.slice(1);

    const sp = bp.section10Particulars;
    // Semantic keywords first (so OCR letter confusion on "j." vs "l." never misroutes), prefix letters last
    const semanticMatchers: Array<{ key: keyof Section10Particulars; re: RegExp }> = [
      { key: 'nameOfClient', re: /Name\s+o[fl]'?\s+th|ol'th[ec]\s+CIicnt/i },
      { key: 'uniqueCustomerId', re: /Customer\s*ID|Lrstonrcry/i },
      { key: 'awcdMsdAccount', re: /AWCD|WCD|MSD/i },
      { key: 'businessTypeNature', re: /Business\s+Type/i },
      { key: 'businessManagementType', re: /Management/i },
      { key: 'clientType', re: /Client\s+Type|^f\.?\s+CIicnt/i },
      { key: 'keyPersonPhone', re: /Phone|I\)honc|Key\s+Person/i },
      { key: 'officeBusinessAddress', re: /Office\/Business|IJtrsirtesr/i },
      { key: 'factoryAddress', re: /Factory|clclrcss/i },
      { key: 'address', re: /Aclch'css|\bAddress\b/i },
      { key: 'purposeOfInvestment', re: /Purpose\s+of/i },
      { key: 'limitOfInvestment', re: /Limit\s+of|l-inrit\s+ol/i },
      { key: 'modeOfInvestment', re: /Mode\s+of|Nloclc\s+o1/i },
      { key: 'lastVisitDate', re: /Last\s+Visit|I-ast\s+Visit/i },
      { key: 'stockPosition', re: /Stock\s+Position/i },
      { key: 'fixedAsset', re: /Fixed\s+Asset|irccl\s+Assct/i },
      { key: 'businessNetWorth', re: /Business\s+Net|lJLrsincss\s+Nct/i },
      { key: 'insuranceCoverage', re: /Insurance|nsurance/i },
      { key: 'tradeLicenseValidity', re: /Trade\s+License|raclc\s+\[.iccnse/i },
      { key: 'otherRegulatoryDocuments', re: /Other\s+Regulatory|llegulatory/i },
      { key: 'dateOfOriginalInvestment', re: /Date\s+of\s+Original|nrrestntcnt/i },
      { key: 'firstDisbursementManager', re: /Disbursement\s+Manager/i },
      { key: 'firstDateOfClassification', re: /First\s+Date\s+o[fl]|Classif/i },
    ];

    const prefixMatchers: Array<{ key: keyof Section10Particulars; re: RegExp }> = [
      { key: 'nameOfClient', re: /^a\.?\s/i },
      { key: 'uniqueCustomerId', re: /^b\.?\s/i },
      { key: 'awcdMsdAccount', re: /^c\.?\s/i },
      { key: 'businessTypeNature', re: /^d\.?\s/i },
      { key: 'businessManagementType', re: /^e\.?\s/i },
      { key: 'clientType', re: /^f\.?\s/i },
      { key: 'officeBusinessAddress', re: /^h\.?$/i },
      { key: 'purposeOfInvestment', re: /^i\.?\s/i },
      { key: 'limitOfInvestment', re: /^j\.?\s/i },
      { key: 'modeOfInvestment', re: /^k\.?\s/i },
      { key: 'lastVisitDate', re: /^l\.?\s/i },
      { key: 'stockPosition', re: /^m\.?\s/i },
      { key: 'fixedAsset', re: /^n\.?\s/i },
      { key: 'businessNetWorth', re: /^o\.?\s/i },
      { key: 'insuranceCoverage', re: /^p\.?\s/i },
      { key: 'tradeLicenseValidity', re: /^q\.?\s/i },
      { key: 'otherRegulatoryDocuments', re: /^r\.?\s/i },
      { key: 'dateOfOriginalInvestment', re: /^s\.?\s/i },
      { key: 'firstDisbursementManager', re: /^t\.?\s/i },
      { key: 'firstDateOfClassification', re: /^u\.?\s/i },
    ];

    for (const pl of partLines) {
      if (isPageHeaderFooterLine(pl.text)) continue;
      const valRight = normalizeCommonOcrWords(itemsInXBand(pl.items, 270, 900));
      if (!valRight) continue;

      const labelLeft = itemsInXBand(pl.items, 55, 270);
      let matchedKey: keyof Section10Particulars | null = null;

      for (const { key, re } of semanticMatchers) {
        if (re.test(labelLeft)) {
          matchedKey = key;
          break;
        }
      }
      if (!matchedKey) {
        for (const { key, re } of prefixMatchers) {
          if (re.test(labelLeft)) {
            matchedKey = key;
            break;
          }
        }
      }

      if (matchedKey && !sp[matchedKey]) {
        let val = valRight;
        if (matchedKey === 'limitOfInvestment') {
          const num = normalizeOcrNumericToken(val);
          if (num) val = /million/i.test(val) ? `${num} Million` : num;
        } else if (matchedKey === 'lastVisitDate') {
          const d = extractDateToken(val);
          if (d) val = d;
        }
        sp[matchedKey] = val;
      }
    }

    // Also parse Directors & Bank Directors tables if non-empty
    if (dirIdx !== -1) {
      const bankDirIdx = sec10.slice.findIndex((l) => /Director\s+of\s+an[yv]?\s+Bank/i.test(l.text));
      const dirEnd = bankDirIdx !== -1 ? bankDirIdx : sec10.slice.length;
      const dRows: Section10DirectorRow[] = [];
      for (const dl of sec10.slice.slice(dirIdx + 2, dirEnd)) {
        if (/Name\s+of\s+the\s+Directors|Age|Education/i.test(dl.text)) continue;
        const row: Section10DirectorRow = {
          name: itemsInXBand(dl.items, 60, 195),
          age: itemsInXBand(dl.items, 195, 242),
          education: itemsInXBand(dl.items, 242, 305),
          designation: itemsInXBand(dl.items, 305, 378),
          sharePct: itemsInXBand(dl.items, 378, 435),
          experience: itemsInXBand(dl.items, 435, 520),
          netWorth: itemsInXBand(dl.items, 520, 900),
        };
        if (Object.values(row).some(Boolean)) dRows.push(row);
      }
      bp.section10Directors = dRows;

      if (bankDirIdx !== -1) {
        const bRows: Section10BankDirectorRow[] = [];
        for (const bl of sec10.slice.slice(bankDirIdx + 2)) {
          if (/Name\s+of\s+the\s+Director|bank\/NBFIs|Percentage|Shares/i.test(bl.text)) continue;
          const row: Section10BankDirectorRow = {
            name: itemsInXBand(bl.items, 60, 260),
            bankNbfi: itemsInXBand(bl.items, 260, 365),
            designation: itemsInXBand(bl.items, 365, 432),
            sharePct: itemsInXBand(bl.items, 432, 515),
            numShares: itemsInXBand(bl.items, 515, 900),
          };
          if (Object.values(row).some(Boolean)) bRows.push(row);
        }
        bp.section10BankDirectors = bRows;
      }
    }
  }

  // 14. Section 11: Financials of The Client
  const sec11 = slicePdfLinesBetween(
    lines,
    /^\s*(?:11|I\s*l)\s*\.\s*F/i,
    /^\s*12\s*\.\s*/i
  );
  if (sec11.slice.length > 0) {
    const hdrText = sec11.slice[0].text;
    if (/\b(?:l?Nil|N\/A)\b/i.test(hdrText)) {
      bp.section11FinancialsNote = /N\/A/i.test(hdrText) ? 'N/A' : 'Nil';
    }
    const fRows: Section11Row[] = [];
    for (const fl of sec11.slice.slice(1)) {
      if (/Ycar|Year|Gross|Sales|Total\s*Asset/i.test(fl.text)) continue;
      if (/^\s*Nil\s*$/i.test(fl.text.trim())) {
        if (!bp.section11FinancialsNote) bp.section11FinancialsNote = 'Nil';
        continue;
      }
      if (/^\s*a\s*$/i.test(fl.text.trim())) continue;
      const row: Section11Row = {
        year: itemsInXBand(fl.items, 60, 140),
        grossSales: itemsInXBand(fl.items, 140, 218),
        grossProfit: itemsInXBand(fl.items, 218, 292),
        netProfit: itemsInXBand(fl.items, 292, 360),
        totalAsset: itemsInXBand(fl.items, 360, 435),
        totalLiabilities: itemsInXBand(fl.items, 435, 520),
        netWorth: itemsInXBand(fl.items, 520, 900),
      };
      if (Object.values(row).some(Boolean)) fRows.push(row);
    }
    bp.section11Rows = fRows;
  }

  // 15. Section 12: Branch Visit Report
  const sec12 = slicePdfLinesBetween(
    lines,
    /^\s*12\s*\.\s*.*Visit\s*Report/i,
    /^\s*13\s*\.\s*/i
  );
  if (sec12.slice.length > 0) {
    const txt12 = cleanText(
      sec12.slice
        .map((l) => l.text)
        .join(' ')
        .replace(/^12\s*\.\s*[^:]+:\s*/i, '')
    );
    bp.section12BranchVisitReport = txt12;
  }

  // 16. Section 15: Recommendation of Branch Investment Committee
  const sec15 = slicePdfLinesBetween(
    lines,
    /^\s*(?:15|l5)\b.*(?:Bran\s*ch|Committee|I\{ecomm)/i,
    /^\s*16\s*\.\s*/i
  );
  if (sec15.slice.length > 0) {
    bp.section15BranchBicRec = cleanText(
      sec15.slice
        .map((l) => l.text)
        .join(' ')
        .replace(/^(?:15|l5)\s*\.?\s*[^:]+:\s*/i, '')
    );
  }

  // 17. Section 16: Branch Recommendation
  const sec16 = slicePdfLinesBetween(
    lines,
    /^\s*16\s*\.\s*Branch\s+Recommendation/i,
    /Submitted\s+for\s+kind\s+decision/i
  );
  if (sec16.slice.length > 0) {
    bp.section16BranchRec = cleanText(
      sec16.slice
        .map((l) => l.text)
        .join(' ')
        .replace(/^16\s*\.\s*Branch\s+Recommendation\s*:?\s*/i, '')
    );
  }

  return bp;
}

/**
 * Rule-based non-AI extractor for Word (.docx) and Excel (.xlsx / .xls / .csv) Branch Proposals.
 * Reads directly from tables, grids, and paragraphs in milliseconds without any AI API call.
 */
export function parseBranchProposalFromStructured(
  grids: Array<Array<Array<string | number | null>>>,
  lines: string[],
  rawText: string
): BranchProposalData {
  const bp = createEmptyBranchProposal();
  if (!lines.length && !grids.length && !rawText.trim()) return bp;

  // Scan paragraphs & line-form rows first
  for (let i = 0; i < lines.length; i++) {
    const l = cleanText(lines[i]);

    if (!bp.referenceNo) {
      const refM = l.match(/Ref(?:\.?\s*No\.?)?\s*:?\s*([A-Z]{2,}[A-Z0-9.-]*(?:\/[A-Z0-9.-]+){2,})/i);
      if (refM) bp.referenceNo = refM[1].trim();
      else {
        const bareRef = l.match(/\b([A-Z]{2,6}\/[A-Z0-9.-]+\/[A-Z0-9.-]+\/\d{4}\/[A-Z0-9-]+)\b/);
        if (bareRef) bp.referenceNo = bareRef[1].trim();
      }
    }

    if (!bp.proposalDate && /\bDate\s*:/i.test(l) && !/circular/i.test(l)) {
      const d = extractDateToken(l.replace(/^.*Date\s*:/i, ''));
      if (d) bp.proposalDate = d;
    }

    if (!bp.branchName) {
      const brM = l.match(/\bBranch(?:\s*Name)?\s*:\s*([^|]+)/i);
      if (brM) bp.branchName = cleanText(brM[1]);
    }

    if (!bp.accountNo) {
      const acM = l.match(/\bA\/C(?:\s*No\.?)?\s*:?\s*([A-Za-z0-9-]{6,})/);
      if (acM) bp.accountNo = cleanText(acM[1]);
    }

    if (!bp.subjectText && /^Subject\s*:/i.test(l)) {
      bp.subjectText = cleanText(l.replace(/^Subject\s*:/i, ''));
    }

    if (!bp.natureOfBusiness && /(?:^|\b)1\s*\.\s*Nature\s+of\s+Business\s*:/i.test(l)) {
      bp.natureOfBusiness = cleanText(l.replace(/^.*Nature\s+of\s+Business\s*:\s*/i, ''));
    }

    if (!bp.previousBoardDecisions && /Previous\s+Board\s+Decision/i.test(l)) {
      bp.previousBoardDecisions = cleanText(l.replace(/^.*Previous\s+Board\s+Decision[^:]*:\s*/i, ''));
    }

    if (!bp.section5StatusNote && /(?:^|\b)5\s*\.\s*Legal\s+Action/i.test(l)) {
      const m = l.match(/\b(N\/A|Nil)\b/i);
      if (m) bp.section5StatusNote = m[1];
    }

    if (!bp.section6StatusNote && /(?:^|\b)6\s*\.\s*Audit/i.test(l)) {
      const m = l.match(/\b(N\/A|Nil)\b/i);
      if (m) bp.section6StatusNote = m[1];
    }

    if (!bp.section3AsOnDate && /(?:^|\b)3\s*\.\s*Liability/i.test(l)) {
      const d = extractDateToken(l);
      if (d) bp.section3AsOnDate = d;
    }

    if (!bp.section4CibDate && /(?:^|\b)4\s*\.\s*CIB/i.test(l)) {
      const d = extractDateToken(l);
      if (d) bp.section4CibDate = d;
    }

    if (!bp.section8SourceOfRepayment && /(?:^|\b)8\s*\.\s*Source\s+of\s+Repayment\s*:/i.test(l)) {
      bp.section8SourceOfRepayment = cleanText(
        l.replace(/^.*Source\s+of\s+Repayment\s*:\s*/i, '')
      );
    }

    if (!bp.section11FinancialsNote && /(?:^|\b)11\s*\.\s*Financials/i.test(l)) {
      const m = l.match(/\b(N\/A|Nil)\b/i);
      if (m) bp.section11FinancialsNote = m[1];
    }

    if (!bp.section12BranchVisitReport && /(?:^|\b)12\s*\.\s*Branch\s+Visit\s+Report\s*:/i.test(l)) {
      bp.section12BranchVisitReport = cleanText(
        l.replace(/^.*Branch\s+Visit\s+Report\s*:\s*/i, '')
      );
    }

    if (!bp.section15BranchBicRec && /(?:^|\b)15\s*\.\s*Recommendation\s+of\s+Branch/i.test(l)) {
      bp.section15BranchBicRec = cleanText(l.replace(/^.*Recommendation\s+of\s+Branch[^:]*:\s*/i, ''));
    }

    if (!bp.section16BranchRec && /(?:^|\b)16\s*\.\s*Branch\s+Recommendation\s*:/i.test(l)) {
      bp.section16BranchRec = cleanText(l.replace(/^.*Branch\s+Recommendation\s*:\s*/i, ''));
    }
  }

  // Multi-line justification capture
  const justStart = lines.findIndex((l) => /Justification\s+of\s+(?:The\s+)?Proposed\s+Facility/i.test(l));
  if (justStart !== -1) {
    const buf: string[] = [];
    for (let i = justStart + 1; i < lines.length; i++) {
      if (/Previous\s+Board\s+Decision|^1\s*\.\s*Nature/i.test(lines[i])) break;
      buf.push(lines[i]);
    }
    bp.justificationText = cleanText(buf.join(' '));
  }

  // Scan every table/sheet grid by recognizing table headers
  const strGrid = (grid: Array<Array<string | number | null>>) =>
    grid.map((row) => row.map((c) => (c === null || c === undefined ? '' : cleanText(String(c)))));

  for (const rawGrid of grids) {
    const grid = strGrid(rawGrid);
    if (grid.length === 0) continue;

    // Key-value label lookup inside any grid (handles Branch Ref, Date, Section 10 Particulars, etc.)
    for (let r = 0; r < grid.length; r++) {
      const row = grid[r];
      const joined = row.join(' ');

      if (!bp.referenceNo) {
        const refM = joined.match(/Ref(?:\.?\s*No\.?)?\s*:?\s*([A-Z]{2,}[A-Z0-9.-]*(?:\/[A-Z0-9.-]+){2,})/i);
        if (refM) bp.referenceNo = refM[1].trim();
      }
      if (!bp.proposalDate && /\bDate\b/i.test(joined) && !/circular|expiry|filing|visit|inquiry/i.test(joined)) {
        const d = extractDateToken(joined);
        if (d) bp.proposalDate = d;
      }

      // Section 10 Particulars row matching ("a. | Name of the Client | : | Value")
      const sp = bp.section10Particulars;
      const spLabels: Array<{ key: keyof Section10Particulars; re: RegExp }> = [
        { key: 'nameOfClient', re: /Name\s+of\s+the\s+Client/i },
        { key: 'uniqueCustomerId', re: /Unique\s+Customer\s*ID/i },
        { key: 'awcdMsdAccount', re: /AWCD\/?MSD/i },
        { key: 'businessTypeNature', re: /Business\s+Type/i },
        { key: 'businessManagementType', re: /Business\s+Management/i },
        { key: 'clientType', re: /Client\s+Type/i },
        { key: 'keyPersonPhone', re: /Key\s+Person/i },
        { key: 'officeBusinessAddress', re: /Office\/Business\s+Address/i },
        { key: 'factoryAddress', re: /Factory\s+Address/i },
        { key: 'address', re: /^Address$/i },
        { key: 'purposeOfInvestment', re: /Purpose\s+of\s+Investment/i },
        { key: 'limitOfInvestment', re: /Limit\s+of\s+Investment/i },
        { key: 'modeOfInvestment', re: /Mode\s+of\s+Investment/i },
        { key: 'lastVisitDate', re: /Last\s+Visit\s+Date/i },
        { key: 'stockPosition', re: /Stock\s+Position/i },
        { key: 'fixedAsset', re: /Fixed\s+Asset/i },
        { key: 'businessNetWorth', re: /Business\s+Net\s+Worth/i },
        { key: 'insuranceCoverage', re: /Insurance\s+Coverage/i },
        { key: 'tradeLicenseValidity', re: /Trade\s+License/i },
        { key: 'otherRegulatoryDocuments', re: /Other\s+Regulatory/i },
        { key: 'dateOfOriginalInvestment', re: /Date\s+of\s+Original\s+Investment/i },
        { key: 'firstDisbursementManager', re: /Disbursement\s+Manager/i },
        { key: 'firstDateOfClassification', re: /First\s+Date\s+of\s+Classification/i },
      ];

      for (const { key, re } of spLabels) {
        const colIdx = row.findIndex((c) => re.test(c));
        if (colIdx !== -1 && !sp[key]) {
          for (let k = row.length - 1; k > colIdx; k--) {
            if (row[k] && row[k] !== ':') {
              sp[key] = row[k];
              break;
            }
          }
        }
      }
    }

    // Table 1: Main settlement table ("Mode" + "Original" + "Proposed Waiver")
    const mainHdr = grid.findIndex(
      (r) => /Mode/i.test(r.join(' ')) && /Particulars/i.test(r.join(' '))
    );
    if (mainHdr !== -1) {
      const hdr0 = grid[mainHdr];
      const hdr1 = grid[mainHdr + 1] ?? [];
      const origCol = hdr1.findIndex((c) => /Original/i.test(c));
      const rsdCol = hdr1.findIndex((c) => /\bRSD\b/i.test(c));
      const remCol = hdr0.findIndex((c) => /Remarks/i.test(c));

      for (let r = mainHdr + 2; r < grid.length; r++) {
        if (!bp.limitOriginal && origCol >= 0) {
          const num = normalizeOcrNumericToken(grid[r][origCol] ?? '');
          if (num) bp.limitOriginal = num;
        }
        if (!bp.limitRSD && rsdCol >= 0 && grid[r][rsdCol] && !/RSD/i.test(grid[r][rsdCol])) {
          bp.limitRSD = grid[r][rsdCol];
        }
        const remTxt = remCol >= 0 ? grid[r][remCol] : grid[r].join(' ');
        const mvM = remTxt.match(/MV\s*:\s*([A-Za-z0-9.,/-]+)/i);
        if (mvM && !bp.securityMV) bp.securityMV = cleanText(mvM[1]);
        const fsvM = remTxt.match(/FSV\s*:\s*([A-Za-z0-9.,/-]+)/i);
        if (fsvM && !bp.securityFSV) bp.securityFSV = cleanText(fsvM[1]);
      }
    }

    // Table 2: Section 2 Liability ("Mode of Investment" + "Validity" + "Recovery")
    const sec2Hdr = grid.findIndex((r) => {
      const j = r.join(' ');
      return /Mode\s+of\s+Investment/i.test(j) && /Validity/i.test(j);
    });
    if (sec2Hdr !== -1) {
      const hdr = grid[sec2Hdr];
      const limC = hdr.findIndex((c) => /\bLimit\b/i.test(c));
      const valC = hdr.findIndex((c) => /Validity/i.test(c));
      const ovdC = hdr.findIndex((c) => /Overdue/i.test(c));
      const recC = hdr.findIndex((c) => /Recovery/i.test(c));
      const remC = hdr.findIndex((c) => /Remarks/i.test(c));

      for (let r = sec2Hdr + 1; r < grid.length; r++) {
        const joined = grid[r].join(' ');
        if (/^Total\b|Non\s*-?\s*Funded/i.test(joined)) break;
        if (!bp.sec2Limit && limC >= 0 && grid[r][limC]) {
          bp.sec2Limit = normalizeOcrNumericToken(grid[r][limC]) || grid[r][limC];
        }
        if (!bp.sec2Validity && valC >= 0 && grid[r][valC]) {
          bp.sec2Validity = grid[r][valC];
        }
        if (!bp.sec2Overdue && ovdC >= 0 && grid[r][ovdC]) {
          bp.sec2Overdue = grid[r][ovdC];
        }
        if (!bp.sec2Recovery && recC >= 0 && grid[r][recC]) {
          bp.sec2Recovery = grid[r][recC];
        }
        if (!bp.sec2Remarks && remC >= 0 && grid[r][remC]) {
          bp.sec2Remarks = grid[r][remC];
        }
      }
    }

    // Table 3: Section 3 ("Name of the Banks/NBFIs" + "CL Status")
    const sec3Hdr = grid.findIndex((r) => {
      const j = r.join(' ');
      return /Banks\/NBFIs/i.test(j) && /CL\s*Status/i.test(j);
    });
    if (sec3Hdr !== -1 && bp.section3Rows.length === 0) {
      for (let r = sec3Hdr + 1; r < grid.length; r++) {
        const row = grid[r];
        if (!row.some(Boolean)) continue;
        bp.section3Rows.push({
          sl: row[0] ?? '',
          nameOfConcern: row[1] ?? '',
          bankOrNbfi: row[2] ?? '',
          natureOfFacility: row[3] ?? '',
          limit: row[4] ?? '',
          outstanding: row[5] ?? '',
          overdueEol: row[6] ?? '',
          expiryDate: row[7] ?? '',
          clStatus: row[8] ?? '',
        });
      }
    }

    // Table 4: Section 4 CIB Status ("With respective unit" or "Willful Default")
    const sec4Hdr = grid.findIndex((r) => /STATUS/i.test(r.join(' ')) && /respective\s*unit/i.test(r.join(' ')));
    if (sec4Hdr !== -1 && bp.section4Rows.length === 0) {
      const startR = /Borrower/i.test((grid[sec4Hdr + 1] ?? []).join(' ')) ? sec4Hdr + 2 : sec4Hdr + 1;
      for (let r = startR; r < grid.length; r++) {
        const row = grid[r];
        if (!row[0]) continue;
        bp.section4Rows.push({
          status: row[0] ?? '',
          unitBorrower: row[1] ?? '',
          unitGuarantor: row[2] ?? '',
          otherBorrower: row[3] ?? '',
          otherGuarantor: row[4] ?? '',
          remarks: row[5] ?? '',
        });
      }
    }

    // Table 5: Section 5 Legal Action ("Type of Case" + "Case No.")
    const sec5Hdr = grid.findIndex((r) => /Type\s*of\s*Case/i.test(r.join(' ')) && /Case\s*No/i.test(r.join(' ')));
    if (sec5Hdr !== -1 && bp.section5Rows.length === 0) {
      for (let r = sec5Hdr + 1; r < grid.length; r++) {
        const row = grid[r];
        if (!row.some(Boolean)) continue;
        bp.section5Rows.push({
          sl: row[0] ?? '',
          typeOfCase: row[1] ?? '',
          caseNo: row[2] ?? '',
          filingDate: row[3] ?? '',
          amount: row[4] ?? '',
          previousDate: row[5] ?? '',
          nextDate: row[6] ?? '',
          presentStatus: row[7] ?? '',
        });
      }
    }

    // Table 6: Section 6 Audit ("Name of the Authority" + "Objection Status")
    const sec6Hdr = grid.findIndex((r) => /Authority/i.test(r.join(' ')) && /Objection/i.test(r.join(' ')));
    if (sec6Hdr !== -1 && bp.section6Rows.length === 0) {
      for (let r = sec6Hdr + 1; r < grid.length; r++) {
        const row = grid[r];
        if (!row.some(Boolean)) continue;
        bp.section6Rows.push({
          sl: row[0] ?? '',
          authority: row[1] ?? '',
          lastInspectionDate: row[2] ?? '',
          objectionStatus: row[3] ?? '',
          objections: row[4] ?? '',
          branchReply: row[5] ?? '',
        });
      }
    }

    // Table 7a: Primary Security
    const sec7aHdr = grid.findIndex(
      (r) => /Particulars/i.test(r.join(' ')) && /Value/i.test(r.join(' ')) && r.length <= 4
    );
    if (sec7aHdr !== -1 && bp.section7PrimaryRows.length === 0) {
      const pRows: Section7PrimaryRow[] = [];
      for (let r = sec7aHdr + 1; r < grid.length; r++) {
        const row = grid[r];
        if (/Total\s+value\s+of\s+primary/i.test(row.join(' '))) {
          bp.section7PrimaryTotal = row[row.length - 1] ?? '';
          break;
        }
        if (row[0]) {
          pRows.push({ particulars: row[0], value: row[1] ?? '' });
        }
      }
      bp.section7PrimaryRows = pRows;
    }

    // Table 7b: Valuation
    const sec7ValHdr = grid.findIndex((r) => /Value\s+of\s+the\s+mortgaged/i.test(r.join(' ')));
    if (sec7ValHdr !== -1 && bp.section7ValuationRows.length === 0) {
      const mvIdx = grid.findIndex((r, idx) => idx >= sec7ValHdr && /\bMV\b/.test(r.join(' ')) && /\bFSV\b/.test(r.join(' ')));
      if (mvIdx !== -1) {
        for (let r = mvIdx + 1; r < grid.length; r++) {
          const row = grid[r];
          if (/Coverage/i.test(row.join(' '))) {
            const m = row.join(' ').match(/(\d+(?:\.\d+)?\s*%)/);
            if (m) bp.section7CoveragePct = m[1].replace(/\s+/g, '');
            break;
          }
          if (!row[0]) continue;
          bp.section7ValuationRows.push({
            schedule: row[0] ?? '',
            origBranchMv: row[1] ?? '',
            origBranchFsv: row[2] ?? '',
            origSurvMv: row[3] ?? '',
            origSurvFsv: row[4] ?? '',
            presBranchMv: row[5] ?? '',
            presBranchFsv: row[6] ?? '',
            presSurvMv: row[7] ?? '',
            presSurvFsv: row[8] ?? '',
            presAvMv: row[9] ?? '',
            presAvFsv: row[10] ?? '',
          });
        }
      }
    }

    // Table 9: Sequence of HO Sanctions
    const sec9Hdr = grid.findIndex((r) => /HO\s+Sanction\s+Reference/i.test(r.join(' ')));
    if (sec9Hdr !== -1 && bp.section9Rows.length === 0) {
      for (let r = sec9Hdr + 1; r < grid.length; r++) {
        const row = grid[r];
        if (!row.some(Boolean)) continue;
        bp.section9Rows.push({
          sl: row[0] ?? '',
          referenceNo: row[1] ?? '',
          date: row[2] ?? '',
          limit: row[3] ?? '',
          remarks: row[4] ?? '',
        });
      }
    }

    // Table 10b: Directors
    const dirHdr = grid.findIndex((r) => /Name\s+of\s+the\s+Directors/i.test(r.join(' ')) && /Share\s*\(%\)/i.test(r.join(' ')));
    if (dirHdr !== -1 && bp.section10Directors.length === 0) {
      for (let r = dirHdr + 1; r < grid.length; r++) {
        const row = grid[r];
        if (!row.some(Boolean)) continue;
        bp.section10Directors.push({
          name: row[0] ?? '',
          age: row[1] ?? '',
          education: row[2] ?? '',
          designation: row[3] ?? '',
          sharePct: row[4] ?? '',
          experience: row[5] ?? '',
          netWorth: row[6] ?? '',
        });
      }
    }

    // Table 10c: Bank Directors
    const bDirHdr = grid.findIndex((r) => /bank\/NBFIs/i.test(r.join(' ')) && /Number\s+of\s+Shares/i.test(r.join(' ')));
    if (bDirHdr !== -1 && bp.section10BankDirectors.length === 0) {
      for (let r = bDirHdr + 1; r < grid.length; r++) {
        const row = grid[r];
        if (!row.some(Boolean)) continue;
        bp.section10BankDirectors.push({
          name: row[0] ?? '',
          bankNbfi: row[1] ?? '',
          designation: row[2] ?? '',
          sharePct: row[3] ?? '',
          numShares: row[4] ?? '',
        });
      }
    }

    // Table 11: Financials
    const finHdr = grid.findIndex((r) => /Gross\s+Sales/i.test(r.join(' ')) && /Net\s+Profit/i.test(r.join(' ')));
    if (finHdr !== -1 && bp.section11Rows.length === 0) {
      for (let r = finHdr + 1; r < grid.length; r++) {
        const row = grid[r];
        if (!row.some(Boolean)) continue;
        if (/^Nil$/i.test(row.filter(Boolean).join(''))) {
          bp.section11FinancialsNote = 'Nil';
          continue;
        }
        bp.section11Rows.push({
          year: row[0] ?? '',
          grossSales: row[1] ?? '',
          grossProfit: row[2] ?? '',
          netProfit: row[3] ?? '',
          totalAsset: row[4] ?? '',
          totalLiabilities: row[5] ?? '',
          netWorth: row[6] ?? '',
        });
      }
    }
  }

  return bp;
}
