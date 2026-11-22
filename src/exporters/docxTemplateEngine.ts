import JSZip from 'jszip';
import type {
  MemoDocType,
  MemoTemplateRecord,
  ResolvedICMemo,
  ResolvedStageMemo,
  TemplatePreviewBlock,
} from '../types/memo';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function directChildren(parent: Element, localName: string): Element[] {
  const out: Element[] = [];
  for (let i = 0; i < parent.childNodes.length; i++) {
    const n = parent.childNodes[i];
    if (n.nodeType === 1 && (n as Element).localName === localName) {
      out.push(n as Element);
    }
  }
  return out;
}

function firstDescendant(parent: Element, localName: string): Element | null {
  for (let i = 0; i < parent.childNodes.length; i++) {
    const n = parent.childNodes[i];
    if (n.nodeType === 1) {
      const el = n as Element;
      if (el.localName === localName) return el;
      const found = firstDescendant(el, localName);
      if (found) return found;
    }
  }
  return null;
}

/** Collect <w:t> nodes belonging directly to this <w:p> (ignoring nested <w:p> inside textboxes) */
function collectDirectTextNodes(pEl: Element): Element[] {
  const nodes: Element[] = [];
  const walk = (el: Element) => {
    for (let i = 0; i < el.childNodes.length; i++) {
      const child = el.childNodes[i];
      if (child.nodeType !== 1) continue;
      const cEl = child as Element;
      if (cEl.localName === 'p') continue; // Do not cross into nested textbox paragraphs
      if (cEl.localName === 't') {
        nodes.push(cEl);
      } else if (cEl.localName !== 'delText' && cEl.localName !== 'instrText') {
        walk(cEl);
      }
    }
  };
  walk(pEl);
  return nodes;
}

function getParagraphPlainText(pEl: Element): string {
  const parts: string[] = [];
  const walk = (el: Element) => {
    for (let i = 0; i < el.childNodes.length; i++) {
      const child = el.childNodes[i];
      if (child.nodeType !== 1) continue;
      const cEl = child as Element;
      if (cEl.localName === 'p') continue;
      if (cEl.localName === 't') {
        parts.push(cEl.textContent ?? '');
      } else if (cEl.localName === 'br' || cEl.localName === 'cr') {
        parts.push('\n');
      } else if (cEl.localName === 'tab') {
        parts.push(' ');
      } else if (cEl.localName !== 'delText' && cEl.localName !== 'instrText') {
        walk(cEl);
      }
    }
  };
  walk(pEl);
  return parts.join('');
}

function getCellPlainText(tcEl: Element): string {
  return directChildren(tcEl, 'p')
    .map((p) => getParagraphPlainText(p).trim())
    .filter(Boolean)
    .join('\n');
}

function setTextNodeValue(tEl: Element, value: string) {
  tEl.textContent = value;
  if (value.startsWith(' ') || value.endsWith(' ') || value.includes('  ')) {
    tEl.setAttribute('xml:space', 'preserve');
  }
}

/**
 * Replaces a substring [start, end) across potentially split <w:t> nodes in a <w:p>,
 * preserving the <w:rPr> of the first overlapping run and handling multi-line values via <w:br/>.
 */
function replaceSpanInParagraph(
  xmlDoc: Document,
  pEl: Element,
  start: number,
  end: number,
  replacement: string
) {
  const tNodes = collectDirectTextNodes(pEl);
  if (tNodes.length === 0) return;

  interface Span {
    node: Element;
    start: number;
    end: number;
    text: string;
  }

  const spans: Span[] = [];
  let cursor = 0;
  for (const node of tNodes) {
    const text = node.textContent ?? '';
    spans.push({ node, start: cursor, end: cursor + text.length, text });
    cursor += text.length;
  }

  const overlapping = spans.filter((s) => s.end > start && s.start < end);
  if (overlapping.length === 0) return;

  const first = overlapping[0];
  const last = overlapping[overlapping.length - 1];
  const prefix = first.text.slice(0, Math.max(0, start - first.start));
  const suffix = last.text.slice(Math.max(0, end - last.start));

  for (let i = 1; i < overlapping.length; i++) {
    setTextNodeValue(overlapping[i].node, '');
  }

  if (!replacement.includes('\n')) {
    setTextNodeValue(first.node, prefix + replacement + suffix);
    return;
  }

  // Multi-line replacement inside a single <w:r>: split by \n and insert <w:br/>
  const parentRun = first.node.parentElement;
  if (!parentRun || parentRun.localName !== 'r') {
    setTextNodeValue(first.node, prefix + replacement.replace(/\n+/g, ' ') + suffix);
    return;
  }

  const lines = replacement.split('\n');
  setTextNodeValue(first.node, prefix + lines[0]);

  let anchor: Node = first.node;
  for (let i = 1; i < lines.length; i++) {
    const br = xmlDoc.createElementNS(W_NS, 'w:br');
    parentRun.insertBefore(br, anchor.nextSibling);
    anchor = br;

    const t = xmlDoc.createElementNS(W_NS, 'w:t');
    const lineVal = i === lines.length - 1 ? lines[i] + suffix : lines[i];
    setTextNodeValue(t, lineVal);
    parentRun.insertBefore(t, anchor.nextSibling);
    anchor = t;
  }
}

/**
 * Replaces the full visible text of a paragraph while keeping <w:pPr> and first run's <w:rPr> intact.
 */
function setParagraphTextPreservingStyle(
  xmlDoc: Document,
  pEl: Element,
  newText: string
) {
  const pPr = directChildren(pEl, 'pPr')[0] ?? null;
  const firstRun = directChildren(pEl, 'r')[0] ?? null;
  const firstRPr = firstRun ? firstDescendant(firstRun, 'rPr') : null;
  const pRPr = pPr ? firstDescendant(pPr, 'rPr') : null;
  const rPrTemplate = firstRPr
    ? (firstRPr.cloneNode(true) as Element)
    : pRPr
      ? (pRPr.cloneNode(true) as Element)
      : null;

  // Keep only pPr, remove existing runs/text content
  const toRemove: ChildNode[] = [];
  for (let i = 0; i < pEl.childNodes.length; i++) {
    const child = pEl.childNodes[i];
    if (child !== pPr) {
      toRemove.push(child);
    }
  }
  for (const child of toRemove) {
    pEl.removeChild(child);
  }

  const rEl = xmlDoc.createElementNS(W_NS, 'w:r');
  if (rPrTemplate) {
    rEl.appendChild(rPrTemplate);
  }

  const lines = newText.replace(/\r\n/g, '\n').split('\n');
  for (let i = 0; i < lines.length; i++) {
    if (i > 0) {
      rEl.appendChild(xmlDoc.createElementNS(W_NS, 'w:br'));
    }
    const tEl = xmlDoc.createElementNS(W_NS, 'w:t');
    setTextNodeValue(tEl, lines[i]);
    rEl.appendChild(tEl);
  }
  pEl.appendChild(rEl);
}

/**
 * Sets text inside a <w:tc> table cell while strictly preserving <w:tcPr>, <w:pPr>, and <w:rPr>.
 */
function setCellTextPreservingFormat(
  xmlDoc: Document,
  tcEl: Element,
  newText: string
) {
  const paras = directChildren(tcEl, 'p');
  if (paras.length === 0) {
    const pEl = xmlDoc.createElementNS(W_NS, 'w:p');
    tcEl.appendChild(pEl);
    setParagraphTextPreservingStyle(xmlDoc, pEl, newText);
    return;
  }

  // Keep first paragraph's exact <w:pPr> and first run's <w:rPr>
  const firstP = paras[0];
  setParagraphTextPreservingStyle(xmlDoc, firstP, newText);

  // Remove any leftover extra paragraphs inside the cell so multi-line replacements don't collide
  for (let i = 1; i < paras.length; i++) {
    tcEl.removeChild(paras[i]);
  }
}

/**
 * Builds a clean non-bold, non-underlined <w:rPr> inheriting font/size/color from an existing paragraph.
 */
function buildCleanValueRPr(pEl: Element): Element | null {
  const runs = directChildren(pEl, 'r');
  const pPr = directChildren(pEl, 'pPr')[0] ?? null;
  // Prefer the last run's rPr, else first run's rPr, else pPr/rPr
  const lastRun = runs[runs.length - 1];
  const firstRun = runs[0];
  const srcRPr =
    (lastRun ? firstDescendant(lastRun, 'rPr') : null) ??
    (firstRun ? firstDescendant(firstRun, 'rPr') : null) ??
    (pPr ? firstDescendant(pPr, 'rPr') : null);

  if (!srcRPr) return null;
  const cloned = srcRPr.cloneNode(true) as Element;
  for (const tag of ['b', 'bCs', 'u', 'i', 'iCs']) {
    const el = firstDescendant(cloned, tag);
    if (el) el.parentNode?.removeChild(el);
  }
  return cloned;
}

/**
 * Replaces "Label: <old>" inside a paragraph while preserving the bold/underlined label runs,
 * and handles both existing tails AND empty tails cleanly (e.g. "12. Branch Visit Report:").
 */
function setParagraphValueAfterColon(
  xmlDoc: Document,
  pEl: Element,
  newValue: string
): boolean {
  const full = collectDirectTextNodes(pEl)
    .map((t) => t.textContent ?? '')
    .join('');
  const colonIdx = full.indexOf(':');
  if (colonIdx === -1) return false;

  const afterIdx = colonIdx + 1;
  const tail = full.slice(afterIdx);

  if (tail.trim().length > 0) {
    const leadingSpaces = tail.match(/^\s*/)?.[0] ?? ' ';
    const replaceStart = afterIdx + leadingSpaces.length;
    replaceSpanInParagraph(xmlDoc, pEl, replaceStart, full.length, newValue);
    return true;
  }

  // Empty tail after colon: append a clean non-bold/non-underlined run preserving font & size
  const rEl = xmlDoc.createElementNS(W_NS, 'w:r');
  const cleanRPr = buildCleanValueRPr(pEl);
  if (cleanRPr) {
    rEl.appendChild(cleanRPr);
  }

  const lines = newValue.replace(/\r\n/g, '\n').split('\n');
  const needsSpace = tail.length === 0;
  for (let i = 0; i < lines.length; i++) {
    if (i > 0) {
      rEl.appendChild(xmlDoc.createElementNS(W_NS, 'w:br'));
    }
    const tEl = xmlDoc.createElementNS(W_NS, 'w:t');
    const prefix = i === 0 && needsSpace ? ' ' : '';
    setTextNodeValue(tEl, prefix + lines[i]);
    rEl.appendChild(tEl);
  }
  pEl.appendChild(rEl);
  return true;
}

export function buildTemplateTokenMap(
  memo: ResolvedICMemo,
  stageMemo?: ResolvedStageMemo | null
): {
  requiredTokens: Record<string, string>;
  optionalTokens: Record<string, string>;
} {
  const ph = memo.placeholders;
  const bp = memo.branchProposal;
  const clientName = bp?.section10Particulars?.nameOfClient?.trim() || ph['#2#'] || '';

  const requiredTokens: Record<string, string> = {};
  for (const [code, val] of Object.entries(ph)) {
    requiredTokens[code] = val && val.trim() !== '' ? val.trim() : code;
  }

  const si = stageMemo?.stageInputs;
  const ss = stageMemo?.stageSections;

  const semanticMap: Record<string, string> = {
    proposalDate: ph['#1#'] || '',
    dated: ph['#1#'] || '',
    memoDate: ph['#1#'] || '',
    accountNo: ph['#2#'] || '',
    accountNumber: ph['#2#'] || '',
    acNo: ph['#2#'] || '',
    clientName,
    nameOfClient: clientName,
    branchName: ph['#3#'] || '',
    branch: ph['#3#'] || '',
    waiverAmount: ph['#4#'] || '',
    waiver: ph['#4#'] || '',
    mode: ph['#5#'] || '',
    modeOfInvestment: ph['#5#'] || '',
    productName: ph['#5#'] || '',
    limitOriginal: ph['#6#'] || '',
    originalLimit: ph['#6#'] || '',
    limit: ph['#6#'] || '',
    principal: ph['#7#'] || '',
    principalOutstanding: ph['#7#'] || '',
    charges: ph['#8#'] || '',
    chargedProfit: ph['#10#'] || '',
    profitOut: ph['#10#'] || '',
    compensation: ph['#11#'] || '',
    unchargedCompensation: ph['#12#'] || '',
    unappliedProfit: ph['#12#'] || '',
    outstanding: ph['#25#'] || '',
    netOutstanding: ph['#25#'] || '',
    totalOutstanding: ph['#25#'] || '',
    settlementAmount: ph['#27#'] || '',
    exitAmount: ph['#27#'] || '',
    settlementPeriod: ph['#28#'] || '',
    settlementDeadline: ph['#28#'] || '',
    costOfFund: ph['#31#'] || '',
    cof: ph['#31#'] || '',
    breakupTotal: ph['#32#'] || '',
    securityMV: ph['#33#'] || '',
    mv: ph['#33#'] || '',
    securityFSV: ph['#34#'] || '',
    fsv: ph['#34#'] || '',
    downPayment: ph['#35#'] || '',
    downPaymentAccount: ph['#36#'] || '',
    totalProfitComp: ph['#42#'] || '',
    maxWaiver: ph['#44#'] || '',
    waiverRecClientBranch: ph['#45#'] || '',
    waiverRecHO: ph['#47#'] || '',
    justification: memo.justificationLines || '',
    natureOfBusiness: ph['#54#'] || '',
    sec2Limit: ph['#56#'] || '',
    validity: ph['#58#'] || '',
    overdue: ph['#59#'] || '',
    recovery: ph['#60#'] || '',
    cibDate: ph['#62#'] || '',
    refNo: ph['#64#'] || '',
    referenceNo: ph['#64#'] || '',
    branchProposalDate: ph['#65#'] || '',
    osDate: ph['#68#'] || '',
    printDate: ph['#68#'] || '',
    cis: ph['#69#'] || '',
    customerId: ph['#69#'] || '',
    sourceOfRepayment: memo.sourceOfRepayment || '',
    sec8SourceOfRepayment: memo.sourceOfRepayment || '',
    clientPrayer: memo.clientPrayer || '',
    contentsOfClientsPrayer: memo.clientPrayer || '',
    branchVisit: memo.branchVisitReport || '',
    branchVisitReport: memo.branchVisitReport || '',
    branchRecommendation: memo.sec13BranchRecommendation || '',
    sec13BranchRecommendation: memo.sec13BranchRecommendation || '',
    iccdRecommendation: memo.sec14IccdRecommendation || '',
    sec14IccdRecommendation: memo.sec14IccdRecommendation || '',
    samdObservation: memo.samdObservation || '',
    samd: memo.samdObservation || '',
    nonFundedNote: memo.sec2NonFundedNote || '',
    sec7OthersNote: memo.sec7OthersNote || '',
    icMemoNumber: si?.icMemoNumber || '',
    icMemoNo: si?.icMemoNumber || '',
    icMemoDate: si?.icMemoDate || '',
    mecMemoNumber: si?.mecMemoNumber || '',
    mecMemoNo: si?.mecMemoNumber || '',
    mecMemoDate: si?.mecMemoDate || '',
    boardMemoNumber: si?.boardMemoNumber || '',
    boardMemoNo: si?.boardMemoNumber || '',
    boardMemoDate: si?.boardMemoDate || '',
    sanctionRefNumber: si?.sanctionRefNumber || '',
    sanctionNo: si?.sanctionRefNumber || '',
    sanctionDate: si?.sanctionDate || '',
    mecIcRecommendation: ss?.mecIcRecommendation || '',
    boardMecRecommendation: ss?.boardMecRecommendation || '',
    sanctionCrossVerifiedSummary: ss?.sanctionCrossVerifiedSummary || '',
  };

  const optionalTokens: Record<string, string> = {};
  for (const [key, val] of Object.entries(semanticMap)) {
    optionalTokens[`#${key}#`] = val;
    optionalTokens[`{{${key}}}`] = val;
    optionalTokens[`[${key}]`] = val;
  }

  return { requiredTokens, optionalTokens };
}

function replaceTokensInParagraph(
  xmlDoc: Document,
  pEl: Element,
  tokenMap: Record<string, string>
): Set<string> {
  const replacedKeys = new Set<string>();
  const keys = Object.keys(tokenMap).sort((a, b) => b.length - a.length);
  if (keys.length === 0) return replacedKeys;

  // Iterative scan so multiple tokens in one paragraph are all replaced cleanly
  let safety = 0;
  while (safety < 40) {
    safety++;
    const fullText = collectDirectTextNodes(pEl)
      .map((t) => t.textContent ?? '')
      .join('');
    if (!fullText) break;

    let bestIdx = -1;
    let bestKey = '';

    for (const key of keys) {
      const idx = fullText.indexOf(key);
      if (idx !== -1 && (bestIdx === -1 || idx < bestIdx)) {
        // Avoid infinite loop if replacement equals key (e.g., unfilled #code#)
        if (tokenMap[key] !== key) {
          bestIdx = idx;
          bestKey = key;
        }
      }
    }

    if (bestIdx === -1) break;
    replaceSpanInParagraph(xmlDoc, pEl, bestIdx, bestIdx + bestKey.length, tokenMap[bestKey]);
    replacedKeys.add(bestKey);
  }

  return replacedKeys;
}

function hasAnyKnownToken(text: string, tokenKeys: string[]): boolean {
  return tokenKeys.some((k) => text.includes(k));
}

function nonEmpty(v: string | undefined): string | null {
  if (!v) return null;
  const trimmed = v.trim();
  return trimmed ? trimmed : null;
}

function paraReplacedAnyOf(
  replacedTokensMap: Map<Element, Set<string>>,
  pEl: Element,
  candidateKeys: string[]
): boolean {
  const set = replacedTokensMap.get(pEl);
  if (!set || set.size === 0) return false;
  return candidateKeys.some((k) => set.has(k));
}

function cellReplacedAnyToken(
  replacedTokensMap: Map<Element, Set<string>>,
  tcEl: Element
): boolean {
  return directChildren(tcEl, 'p').some((p) => {
    const s = replacedTokensMap.get(p);
    return Boolean(s && s.size > 0);
  });
}

function isSectionBoundaryText(text: string): boolean {
  const t = text.replace(/\s+/g, ' ').trim();
  if (!t) return false;
  if (
    /^(?:Justification\s+of\s+(?:the\s+)?Proposed\s+Facility|Previous\s+Board\s+Decision|Subject\s*:|Key\s+Points\s*:|Note\s*:\s*Information\s+Regarding\s+Down)/i.test(
      t
    )
  ) {
    return true;
  }
  if (/^(?:1[0-7]|[1-9])\s*[.)]\s*[A-Za-z]/.test(t)) return true;
  if (
    /^(?:Nature\s+of\s+Business|Liability\s+Position\s+of|Liability\s+of\s+the\s+Client|CIB\s+Status|Legal\s+Action\s+Status|Audit,\s*Compliance|Security\s*:|Source\s+of\s+Repayment|Sequence\s+of\s+Head\s+Office|Brief\s+Particulars\s+of\s+The\s+Client|Financials\s+of\s+The\s+Client|Contents\s+of\s+(?:the\s+)?Client[’']?s\s+Prayer|Branch\s+Visit\s+Report|Recommendation\s+of\s+(?:the\s+)?(?:Investment\s+Committee|Management\s+Executive\s+Committee|Branch)|AI\s+Cross-Verified\s+Facility|Branch\s+Recommendation|Recommendation\s+of\s+(?:Internal\s+Control|IC\s*&?\s*CD|ICCD)|Observations?\s*(?:and|&|\/)\s*Recommendations?\s*of\s*SAMD|SAMD\s*Observations?|Particulars\s+of\s+(?:the\s+)?(?:Proposed\s+)?Facility|Special\s+Term\(s\))/i.test(
      t
    )
  ) {
    return true;
  }
  if (/^(?:[bc]\.\s*(?:Investment\s+was\s+not\s+created|The\s+client\s+is\s+not\s+a\s+willful)|c\.\s*Others\s*:?)/i.test(t)) {
    return true;
  }
  if (/^(?:Particular\s+of\s+the\s+Proprietor|Remarks\s*:\s*Director\s+of\s+any\s+Bank)/i.test(t)) {
    return true;
  }
  if (
    /^(?:The\s+proposal\s+is\s+hereby\s+placed|Submitted\s+for\s+kind\s+decision|Union\s+Bank\s+PLC\s+Unit|Special\s+Asset\s+Management\s+Division|\([A-Z][a-z]+\s+[A-Z]|Deputy\s+Managing\s+Director|Additional\s+Managing\s+Director|Managing\s+Director\b)/i.test(
      t
    )
  ) {
    return true;
  }
  return false;
}

/**
 * Collects contiguous sibling <w:p> elements directly following `headingP` inside the same parent
 * container, stopping at any table (<w:tbl>) or any subsequent section/closing boundary.
 */
function collectContiguousSubParagraphs(headingP: Element, stopAtSubBullets = false): Element[] {
  const out: Element[] = [];
  let node: ChildNode | null = headingP.nextSibling;
  while (node) {
    if (node.nodeType === 1) {
      const el = node as Element;
      if (el.localName === 'tbl' || firstDescendant(el, 'tbl')) {
        break;
      }
      if (el.localName === 'p') {
        const subTxt = getParagraphPlainText(el).trim();
        if (isSectionBoundaryText(subTxt)) break;
        if (stopAtSubBullets && /^[bc]\.\s+/i.test(subTxt)) break;
        out.push(el);
      } else if (el.localName === 'sdt') {
        const sdtP = firstDescendant(el, 'p');
        if (sdtP) {
          const subTxt = getParagraphPlainText(sdtP).trim();
          if (isSectionBoundaryText(subTxt)) break;
          if (stopAtSubBullets && /^[bc]\.\s+/i.test(subTxt)) break;
          out.push(sdtP);
        }
      }
    }
    node = node.nextSibling;
  }
  return out;
}

/**
 * Inserts/updates body sub-paragraphs right under `headingP`, cloning existing sub-paragraph styling
 * when available or building clean non-bold/non-underlined runs inheriting `headingP` font/size.
 */
function populateSubParagraphLines(
  xmlDoc: Document,
  headingP: Element,
  existingSub: Element[],
  lines: string[],
  defaultIndent = '180'
) {
  if (lines.length === 0) return;

  const styledSub =
    existingSub.find((p) => getParagraphPlainText(p).trim() !== '') ?? existingSub[0] ?? null;

  const createStyledPara = (lineText: string): Element => {
    if (styledSub) {
      const cloned = styledSub.cloneNode(true) as Element;
      setParagraphTextPreservingStyle(xmlDoc, cloned, lineText);
      return cloned;
    }
    const pNew = xmlDoc.createElementNS(W_NS, 'w:p');
    const pPr = xmlDoc.createElementNS(W_NS, 'w:pPr');
    if (defaultIndent) {
      const ind = xmlDoc.createElementNS(W_NS, 'w:ind');
      ind.setAttribute('w:left', defaultIndent);
      pPr.appendChild(ind);
    }
    const sp = xmlDoc.createElementNS(W_NS, 'w:spacing');
    sp.setAttribute('w:before', '15');
    sp.setAttribute('w:after', '15');
    pPr.appendChild(sp);
    pNew.appendChild(pPr);

    const rEl = xmlDoc.createElementNS(W_NS, 'w:r');
    const cleanRPr = buildCleanValueRPr(headingP);
    if (cleanRPr) {
      rEl.appendChild(cleanRPr);
    }
    const tEl = xmlDoc.createElementNS(W_NS, 'w:t');
    setTextNodeValue(tEl, lineText);
    rEl.appendChild(tEl);
    pNew.appendChild(rEl);
    return pNew;
  };

  if (existingSub.length > 0) {
    setParagraphTextPreservingStyle(xmlDoc, existingSub[0], lines[0]);
    let insertAfter: Element = existingSub[0];
    for (let k = 1; k < lines.length; k++) {
      if (k < existingSub.length) {
        setParagraphTextPreservingStyle(xmlDoc, existingSub[k], lines[k]);
        insertAfter = existingSub[k];
      } else {
        const cloned = createStyledPara(lines[k]);
        insertAfter.parentNode?.insertBefore(cloned, insertAfter.nextSibling);
        insertAfter = cloned;
      }
    }
    for (let k = lines.length; k < existingSub.length; k++) {
      existingSub[k].parentNode?.removeChild(existingSub[k]);
    }
  } else {
    let insertAfter: Element = headingP;
    for (const line of lines) {
      const pNew = createStyledPara(line);
      insertAfter.parentNode?.insertBefore(pNew, insertAfter.nextSibling);
      insertAfter = pNew;
    }
  }
}

/**
 * Populates a narrative heading paragraph that may either hold its value inline after ":"
 * OR place its body in sub-paragraph slot(s) below the heading (or adjacent/next-row table cell).
 */
function populateNarrativeHeading(
  xmlDoc: Document,
  pEl: Element,
  valueText: string,
  preferSubParagraphs = false
) {
  const fullText = getParagraphPlainText(pEl).trim();
  const colonIdx = fullText.indexOf(':');
  const existingTail = colonIdx !== -1 ? fullText.slice(colonIdx + 1).trim() : '';
  const subParas = collectContiguousSubParagraphs(pEl);
  const lines = valueText
    .replace(/\r\n/g, '\n')
    .split('\n')
    .map((l) => l.trim())
    .filter(Boolean);
  if (lines.length === 0) return;

  const nonEmptySubParas = subParas.filter((sp) => getParagraphPlainText(sp).trim() !== '');

  // Check if pEl sits in a 2+ cell table row as [Label Cell, (":" Cell), Value Cell]
  const tcParent = pEl.parentElement?.localName === 'tc' ? pEl.parentElement : null;
  const trParent = tcParent?.parentElement?.localName === 'tr' ? tcParent.parentElement : null;
  if (tcParent && trParent) {
    const rowCells = directChildren(trParent, 'tc');
    const tcIdx = rowCells.indexOf(tcParent);
    if (
      rowCells.length >= 2 &&
      tcIdx !== -1 &&
      tcIdx < rowCells.length - 1 &&
      !existingTail &&
      nonEmptySubParas.length === 0
    ) {
      let valIdx = tcIdx + 1;
      if (getCellPlainText(rowCells[valIdx]).trim() === ':' && valIdx + 1 < rowCells.length) {
        valIdx += 1;
      }
      const nextCell = rowCells[valIdx];
      const nextText = getCellPlainText(nextCell).trim();
      if (nextText !== ':' && !isSectionBoundaryText(nextText)) {
        setCellTextPreservingFormat(xmlDoc, nextCell, lines.join('\n'));
        return;
      }
    }

    // Also support 2-row box layout: [Row r: Heading in 1 cell] followed by [Row r+1: empty single cell]
    const tblParent = trParent.parentElement?.localName === 'tbl' ? trParent.parentElement : null;
    if (tblParent && rowCells.length === 1 && !existingTail && nonEmptySubParas.length === 0) {
      const tblRows = directChildren(tblParent, 'tr');
      const trIdx = tblRows.indexOf(trParent);
      if (trIdx !== -1 && trIdx + 1 < tblRows.length) {
        const nextRowCells = directChildren(tblRows[trIdx + 1], 'tc');
        if (nextRowCells.length === 1) {
          const nextRowText = getCellPlainText(nextRowCells[0]).trim();
          if (nextRowText === '') {
            setCellTextPreservingFormat(xmlDoc, nextRowCells[0], lines.join('\n'));
            return;
          }
        }
      }
    }
  }

  // If sub-paragraph slots already exist below the heading and there is no inline tail (or multi-line/preferred),
  // write cleanly into sub-paragraphs so template layout & indentation stay intact.
  if (
    subParas.length > 0 &&
    (!existingTail || preferSubParagraphs || lines.length > 1)
  ) {
    if (existingTail && !preferSubParagraphs) {
      if (setParagraphValueAfterColon(xmlDoc, pEl, lines.join('\n'))) {
        for (const sp of subParas) {
          sp.parentNode?.removeChild(sp);
        }
        return;
      }
    }
    populateSubParagraphLines(xmlDoc, pEl, subParas, lines, '120');
    return;
  }

  if (colonIdx !== -1 && !preferSubParagraphs) {
    if (setParagraphValueAfterColon(xmlDoc, pEl, lines.join('\n'))) {
      return;
    }
  }

  populateSubParagraphLines(xmlDoc, pEl, subParas, lines, '120');
}

/**
 * Clones/updates table data rows in [dataStartIdx, dataEndIdx) preserving original <w:trPr>, <w:tcPr>, borders, fonts, and sizes.
 */
function populateTableDataRowsPreservingFormat(
  xmlDoc: Document,
  tbl: Element,
  dataStartIdx: number,
  dataEndIdx: number,
  rowsValues: string[][],
  allowAppend = true,
  clearWhenEmpty = false
) {
  const trs = directChildren(tbl, 'tr');
  const existingDataTrs = trs.slice(dataStartIdx, dataEndIdx);
  if (existingDataTrs.length === 0) return;

  if (rowsValues.length === 0) {
    if (!clearWhenEmpty) return;
    for (const tr of existingDataTrs) {
      const tcs = directChildren(tr, 'tc');
      for (const tc of tcs) {
        setCellTextPreservingFormat(xmlDoc, tc, '');
      }
    }
    return;
  }

  const prototypeTr = existingDataTrs[0];
  let insertAnchor: Element = existingDataTrs[existingDataTrs.length - 1];

  for (let i = 0; i < rowsValues.length; i++) {
    const vals = rowsValues[i];
    let targetTr: Element;
    if (i < existingDataTrs.length) {
      targetTr = existingDataTrs[i];
    } else if (allowAppend) {
      targetTr = prototypeTr.cloneNode(true) as Element;
      tbl.insertBefore(targetTr, insertAnchor.nextSibling);
      insertAnchor = targetTr;
    } else {
      break;
    }

    const tcs = directChildren(targetTr, 'tc');
    for (let c = 0; c < Math.min(tcs.length, vals.length); c++) {
      setCellTextPreservingFormat(xmlDoc, tcs[c], vals[c] ?? '');
    }
  }

  // Clear leftover empty placeholder rows if source has fewer rows than template slots
  for (let i = rowsValues.length; i < existingDataTrs.length; i++) {
    const tcs = directChildren(existingDataTrs[i], 'tc');
    for (const tc of tcs) {
      setCellTextPreservingFormat(xmlDoc, tc, '');
    }
  }
}

function ensureTableFormattingPreserved(tbl: Element, xmlDoc: Document) {
  let tblPr = directChildren(tbl, 'tblPr')[0] ?? null;
  if (!tblPr) {
    tblPr = xmlDoc.createElementNS(W_NS, 'w:tblPr');
    tbl.insertBefore(tblPr, tbl.firstChild);
  }
  if (!directChildren(tblPr, 'tblLayout')[0]) {
    const layout = xmlDoc.createElementNS(W_NS, 'w:tblLayout');
    layout.setAttribute('w:type', 'fixed');
    tblPr.appendChild(layout);
  }
  const trs = directChildren(tbl, 'tr');
  for (const tr of trs) {
    let trPr = directChildren(tr, 'trPr')[0] ?? null;
    if (!trPr) {
      trPr = xmlDoc.createElementNS(W_NS, 'w:trPr');
      tr.insertBefore(trPr, tr.firstChild);
    }
    if (!directChildren(trPr, 'cantSplit')[0]) {
      trPr.appendChild(xmlDoc.createElementNS(W_NS, 'w:cantSplit'));
    }
  }
}

/**
 * Structural & label-based population for uploaded .docx templates that do not use explicit #code# tokens
 * (or for mixed/untokenized sections and tables in a Bank Memo template).
 * Strictly preserves all tables, nested tables, borders, column widths, fonts, sizes, margins, orientation, headers, and footers.
 */
function applyStructuralPopulations(
  xmlDoc: Document,
  memo: ResolvedICMemo,
  replacedTokensMap: Map<Element, Set<string>>,
  stageMemo?: ResolvedStageMemo | null
) {
  const ph = memo.placeholders;
  const bp = memo.branchProposal;
  const clientName = nonEmpty(bp?.section10Particulars?.nameOfClient) || nonEmpty(ph['#2#']);

  const bodyEl = firstDescendant(xmlDoc.documentElement, 'body');
  if (!bodyEl) return;

  // 1. Process all paragraphs in document order (supporting headings in body, SDTs, and 1-cell layout boxes;
  // multi-column key-value tables are handled below unless explicitly detected as a [Heading | Value] row)
  const allParas = Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'p'));
  for (let i = 0; i < allParas.length; i++) {
    const pEl = allParas[i];
    if (!pEl.parentNode) continue;
    const text = getParagraphPlainText(pEl).trim();
    const normText = text.replace(/\s+/g, ' ').trim();
    if (!normText) continue;

    // 1. Nature of Business:
    if (
      /^(?:1\.\s*)?Nature\s+of\s+Business\s*:?/i.test(normText) &&
      nonEmpty(ph['#54#']) &&
      !paraReplacedAnyOf(replacedTokensMap, pEl, ['#54#', '#natureOfBusiness#', '{{natureOfBusiness}}', '[natureOfBusiness]'])
    ) {
      // Avoid touching multi-column Section 10 row if it sits inside a 3/4 col table (Section 10 handles Business Type/Nature)
      const parentTc = pEl.parentElement?.localName === 'tc' ? pEl.parentElement : null;
      const parentTr = parentTc?.parentElement?.localName === 'tr' ? parentTc.parentElement : null;
      const colCount = parentTr ? directChildren(parentTr, 'tc').length : 0;
      if (colCount <= 2) {
        populateNarrativeHeading(xmlDoc, pEl, ph['#54#'], false);
      }
      continue;
    }

    // 2. Liability Position ... as on ...:
    if (
      /^(?:2\.\s*)?Liability\s+Position.*as\s+on\b/i.test(normText) &&
      nonEmpty(ph['#68#']) &&
      !paraReplacedAnyOf(replacedTokensMap, pEl, ['#68#', '#osDate#', '#printDate#', '{{osDate}}'])
    ) {
      const m = text.match(/as\s+on\s+([^:]+):?/i);
      if (m) {
        const start = text.indexOf(m[1], text.toLowerCase().search(/as\s+on/));
        if (start !== -1) {
          replaceSpanInParagraph(xmlDoc, pEl, start, start + m[1].length, ph['#68#']);
        }
      }
      continue;
    }

    // 4. CIB Status: Report dated ... shows:
    if (
      /^(?:4\.\s*)?CIB\s+Status\s*:/i.test(normText) &&
      nonEmpty(ph['#62#']) &&
      !paraReplacedAnyOf(replacedTokensMap, pEl, ['#62#', '#cibDate#', '{{cibDate}}'])
    ) {
      const m = text.match(/Report\s+dated\s+(.+?)\s+shows/i);
      if (m) {
        const start = text.indexOf(m[1]);
        if (start !== -1) {
          replaceSpanInParagraph(xmlDoc, pEl, start, start + m[1].length, ph['#62#']);
        }
      }
      continue;
    }

    // 5. Legal Action Status header note
    if (/^(?:5\.\s*)?Legal\s+Action\s+Status[^:]*:/i.test(normText) && bp?.section5StatusNote !== undefined) {
      if (nonEmpty(bp.section5StatusNote)) {
        setParagraphValueAfterColon(xmlDoc, pEl, bp.section5StatusNote);
      }
      continue;
    }

    // 6. Audit, Compliance header note
    if (/^(?:6\.\s*)?Audit,\s*Compliance[^:]*:/i.test(normText) && bp?.section6StatusNote !== undefined) {
      if (nonEmpty(bp.section6StatusNote)) {
        setParagraphValueAfterColon(xmlDoc, pEl, bp.section6StatusNote);
      }
      continue;
    }

    // 7a / 7b / 7c Notes & Collateral Schedules
    if (/^a\.\s*Primary\s*:/i.test(normText) && nonEmpty(bp?.section7PrimaryNote)) {
      setParagraphValueAfterColon(xmlDoc, pEl, bp!.section7PrimaryNote);
      continue;
    }
    if (/^b\.\s*Collateral\s*:/i.test(normText)) {
      if (nonEmpty(bp?.section7CollateralNote)) {
        setParagraphValueAfterColon(xmlDoc, pEl, bp!.section7CollateralNote);
      }
      if (bp?.section7CollateralSchedules && bp.section7CollateralSchedules.length > 0) {
        const subParas = collectContiguousSubParagraphs(pEl);
        let schedIdx = 0;
        for (const subP of subParas) {
          if (schedIdx >= bp.section7CollateralSchedules.length) break;
          const sTxt = getParagraphPlainText(subP).trim();
          if (/^c\.\s*Others/i.test(sTxt)) break;
          if (/^Schedule[-\s]*\d+/i.test(sTxt)) {
            setParagraphTextPreservingStyle(xmlDoc, subP, bp.section7CollateralSchedules[schedIdx]);
            schedIdx++;
          }
        }
      }
      continue;
    }
    if (/^c\.\s*Others\s*:?/i.test(normText) && nonEmpty(memo.sec7OthersNote)) {
      const subParas = collectContiguousSubParagraphs(pEl);
      const customLines = memo.sec7OthersNote
        .replace(/\r\n/g, '\n')
        .split('\n')
        .map((l) => l.trim())
        .filter(Boolean);
      if (customLines.length > 0) {
        const lastStandardIdx = subParas.reduce((acc, p, idx) => {
          const t = getParagraphPlainText(p).trim();
          return /^(?:a\.\s*Usual Charge|b\.\s*Acceptance of Branch)/i.test(t) ? idx : acc;
        }, -1);
        const targetSub = subParas.slice(lastStandardIdx + 1);
        const anchor = lastStandardIdx >= 0 ? subParas[lastStandardIdx] : pEl;
        populateSubParagraphLines(xmlDoc, anchor, targetSub, customLines, '320');
      }
      continue;
    }

    // 8. Source of Repayment:
    if (
      /^(?:8\.\s*)?Source\s+of\s+Repayment\s*:?/i.test(normText) &&
      nonEmpty(memo.sourceOfRepayment) &&
      !paraReplacedAnyOf(replacedTokensMap, pEl, [
        '#sourceOfRepayment#',
        '#sec8SourceOfRepayment#',
        '{{sourceOfRepayment}}',
        '[sourceOfRepayment]',
      ])
    ) {
      populateNarrativeHeading(xmlDoc, pEl, memo.sourceOfRepayment, false);
      continue;
    }

    // 12. Contents of the Client's Prayer:
    if (
      /^(?:12\s*[.)]\s*)?Contents\s+of\s+(?:the\s+)?Client[’']?s\s+Prayer\s*:?/i.test(normText) &&
      nonEmpty(memo.clientPrayer) &&
      !paraReplacedAnyOf(replacedTokensMap, pEl, [
        '#clientPrayer#',
        '#contentsOfClientsPrayer#',
        '{{clientPrayer}}',
        '{{contentsOfClientsPrayer}}',
        '[clientPrayer]',
      ])
    ) {
      populateNarrativeHeading(xmlDoc, pEl, memo.clientPrayer, false);
      continue;
    }

    // 12. Branch Visit Report:
    if (
      /^(?:12\s*[.)]\s*)?Branch\s+Visit\s+Report\s*:?/i.test(normText) &&
      nonEmpty(memo.branchVisitReport) &&
      !paraReplacedAnyOf(replacedTokensMap, pEl, [
        '#branchVisit#',
        '#branchVisitReport#',
        '{{branchVisit}}',
        '{{branchVisitReport}}',
        '[branchVisit]',
        '[branchVisitReport]',
      ])
    ) {
      populateNarrativeHeading(xmlDoc, pEl, memo.branchVisitReport, false);
      continue;
    }

    // 13. Recommendation of The Branch:
    if (/^(?:13\.\s*)?Recommendation\s+of\s+(?:the\s+)?Branch\b/i.test(normText)) {
      // Update Proposal Ref & Date if present in heading line and not already tokenized
      let updated = text;
      if (
        nonEmpty(ph['#64#']) &&
        !paraReplacedAnyOf(replacedTokensMap, pEl, ['#64#', '#refNo#', '#referenceNo#'])
      ) {
        updated = updated.replace(/(Proposal\s*Ref\s*:\s*)(.+?)(\s+dated\b)/i, `$1${ph['#64#']}$3`);
      }
      if (
        nonEmpty(ph['#65#']) &&
        !paraReplacedAnyOf(replacedTokensMap, pEl, ['#65#', '#branchProposalDate#'])
      ) {
        updated = updated.replace(/(\bdated\s+)(.+)$/i, `$1${ph['#65#']}`);
      }
      if (updated !== text) {
        const colonIdx = updated.indexOf(':');
        if (colonIdx !== -1) {
          setParagraphValueAfterColon(xmlDoc, pEl, updated.slice(colonIdx + 1).trim());
        }
      }

      if (
        nonEmpty(memo.sec13BranchRecommendation) &&
        !paraReplacedAnyOf(replacedTokensMap, pEl, [
          '#branchRecommendation#',
          '#sec13BranchRecommendation#',
          '{{branchRecommendation}}',
          '{{sec13BranchRecommendation}}',
        ])
      ) {
        const recLines = memo.sec13BranchRecommendation
          .replace(/\r\n/g, '\n')
          .split('\n')
          .map((l) => l.trim())
          .filter(Boolean);
        const subParas = collectContiguousSubParagraphs(pEl);
        const hasTokenSub = subParas.some((sp) =>
          paraReplacedAnyOf(replacedTokensMap, sp, [
            '#branchRecommendation#',
            '#sec13BranchRecommendation#',
            '{{branchRecommendation}}',
            '{{sec13BranchRecommendation}}',
          ])
        );
        if (!hasTokenSub && recLines.length > 0) {
          // If heading has "The Branch Recommended the following...", keep heading intact and populate sub-paragraphs below
          if (/The\s+Branch\s+Recommended/i.test(normText) || subParas.length > 0) {
            populateSubParagraphLines(xmlDoc, pEl, subParas, recLines, '180');
          } else {
            populateNarrativeHeading(xmlDoc, pEl, recLines.join('\n'), false);
          }
        }
      }
      continue;
    }

    // 14. Recommendation of Internal Control & Compliance Division / ICCD:
    if (
      /^(?:14\.\s*)?Recommendation\s+of\s+(?:Internal\s+Control|IC\s*&?\s*CD|ICCD)\b/i.test(normText) &&
      nonEmpty(memo.sec14IccdRecommendation) &&
      !paraReplacedAnyOf(replacedTokensMap, pEl, [
        '#iccdRecommendation#',
        '#sec14IccdRecommendation#',
        '{{iccdRecommendation}}',
        '{{sec14IccdRecommendation}}',
      ])
    ) {
      populateNarrativeHeading(xmlDoc, pEl, memo.sec14IccdRecommendation, false);
      continue;
    }

    // Subject: ... waiver of Tk. ...
    if (
      /^Subject\s*:/i.test(normText) &&
      nonEmpty(ph['#4#']) &&
      !paraReplacedAnyOf(replacedTokensMap, pEl, ['#4#', '#waiverAmount#', '#waiver#'])
    ) {
      const m = text.match(/waiver\s+of\s+Tk\.?\s*([^\s,]+(?:\s*million)?)/i);
      if (m) {
        const start = text.indexOf(m[1], text.toLowerCase().search(/waiver\s+of\s+tk/));
        if (start !== -1) {
          replaceSpanInParagraph(xmlDoc, pEl, start, start + m[1].length, ph['#4#']);
        }
      }
      continue;
    }

    // Justification of The Proposed Facility:
    if (/^(?:\d+\.\s*)?Justification\s+of\s+(?:the\s+)?Proposed\s+Facility\s*:?/i.test(normText) && nonEmpty(memo.justificationLines)) {
      const cleanLines = memo.justificationLines
        .replace(/\r\n/g, '\n')
        .split('\n')
        .map((l) => l.trim())
        .filter(Boolean);
      const justAlreadyTokenized = paraReplacedAnyOf(replacedTokensMap, pEl, [
        '#51#',
        '#justification#',
        '{{justification}}',
      ]);
      if (!justAlreadyTokenized && cleanLines.length > 0) {
        const subParas = collectContiguousSubParagraphs(pEl, true);
        const hasTokenSub = subParas.some((sp) =>
          paraReplacedAnyOf(replacedTokensMap, sp, ['#51#', '#justification#', '{{justification}}'])
        );
        if (!hasTokenSub) {
          const aBulletPara = subParas.find((sp) => /^a[\.\)]/i.test(getParagraphPlainText(sp).trim()));
          if (aBulletPara) {
            const aIdx = subParas.indexOf(aBulletPara);
            const aTxt = getParagraphPlainText(aBulletPara).trim();
            const prefix = aTxt.startsWith('a)') ? 'a) ' : 'a. ';
            if (aTxt.length > 2) {
              replaceSpanInParagraph(xmlDoc, aBulletPara, 2, aTxt.length, ` ${cleanLines.join('\n')}`);
            } else {
              setParagraphTextPreservingStyle(xmlDoc, aBulletPara, `${prefix}${cleanLines.join('\n')}`);
            }
            // Remove extra stale continuation lines between "a." and "b."
            for (let k = aIdx + 1; k < subParas.length; k++) {
              subParas[k].parentNode?.removeChild(subParas[k]);
            }
          } else {
            const colonIdx = text.indexOf(':');
            if (subParas.length === 0 && colonIdx !== -1 && text.slice(colonIdx + 1).trim().length > 0) {
              setParagraphValueAfterColon(xmlDoc, pEl, cleanLines.join('\n'));
            } else {
              const formattedLines = cleanLines.map((l, idx) => (idx === 0 ? `a. ${l}` : l));
              populateSubParagraphLines(xmlDoc, pEl, subParas, formattedLines, '120');
            }
          }
        }
      }
      continue;
    }

    // 15. Observation and Recommendation of SAMD:
    if (
      /^(?:15\s*[.)]\s*)?(?:Observations?\s*(?:and|&|\/)\s*Recommendations?\s*of\s*SAMD|SAMD\s*Observations?(?:\s*(?:and|&|\/)\s*Recommendations?)?)\s*:?/i.test(
        normText
      ) &&
      nonEmpty(memo.samdObservation)
    ) {
      const obsLines = memo.samdObservation
        .replace(/\r\n/g, '\n')
        .split('\n')
        .map((l) => l.trim())
        .filter(Boolean);
      const subParas = collectContiguousSubParagraphs(pEl);
      const alreadyTokenized =
        paraReplacedAnyOf(replacedTokensMap, pEl, ['#samdObservation#', '#samd#', '{{samdObservation}}']) ||
        subParas.some((sp) =>
          paraReplacedAnyOf(replacedTokensMap, sp, ['#samdObservation#', '#samd#', '{{samdObservation}}'])
        );
      if (!alreadyTokenized && obsLines.length > 0) {
        const tcParent = pEl.parentElement?.localName === 'tc' ? pEl.parentElement : null;
        const trParent = tcParent?.parentElement?.localName === 'tr' ? tcParent.parentElement : null;
        const tblParent = trParent?.parentElement?.localName === 'tbl' ? trParent.parentElement : null;
        const rowCells = trParent ? directChildren(trParent, 'tc') : [];
        const isTwoRowBox =
          Boolean(tblParent && trParent && rowCells.length === 1) &&
          (() => {
            const tblRows = directChildren(tblParent!, 'tr');
            const trIdx = tblRows.indexOf(trParent!);
            if (trIdx === -1 || trIdx + 1 >= tblRows.length) return false;
            const nextCells = directChildren(tblRows[trIdx + 1], 'tc');
            return nextCells.length === 1 && getCellPlainText(nextCells[0]).trim() === '';
          })();
        if (rowCells.length >= 2 || isTwoRowBox) {
          populateNarrativeHeading(xmlDoc, pEl, obsLines.join('\n'), true);
        } else {
          // Clear any stale inline text after colon if SAMD points live in sub-paragraphs
          const colonIdx = text.indexOf(':');
          if (colonIdx !== -1 && text.slice(colonIdx + 1).trim().length > 0) {
            setParagraphValueAfterColon(xmlDoc, pEl, '');
          }
          populateSubParagraphLines(xmlDoc, pEl, subParas, obsLines, '180');
        }
      }
      continue;
    }
  }

  // 2. Process all tables in document
  const tables = Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'tbl'));
  for (const tbl of tables) {
    ensureTableFormattingPreserved(tbl, xmlDoc);
    const trs = directChildren(tbl, 'tr');
    if (trs.length === 0) continue;

    const matrix = trs.map((tr) =>
      directChildren(tr, 'tc').map((tc) => ({
        tc,
        text: getCellPlainText(tc),
        hasTokens: cellReplacedAnyToken(replacedTokensMap, tc),
      }))
    );

    const tableText = matrix.flat().map((c) => c.text).join(' | ');

    // A. Metadata Table (Dated / Memo no / To / From + Stage Memo Numbers & Dates)
    if (
      (/\bDated\b/i.test(tableText) && /\bMemo\s*no\b/i.test(tableText)) ||
      /\b(?:IC|MEC|Board)\s+Memo\s+(?:Number|No|Date)\b/i.test(tableText) ||
      /\bSanction\s+(?:Ref|No|Date)\b/i.test(tableText)
    ) {
      const si = stageMemo?.stageInputs;
      for (const row of matrix) {
        const label = row[0]?.text.replace(/\s*:\s*$/, '').trim() ?? '';
        const valCell = row[row.length - 1];
        if (!valCell || valCell.hasTokens) continue;
        if (/^IC\s*Memo\s*(?:Number|No\.?)$/i.test(label) && nonEmpty(si?.icMemoNumber)) {
          setCellTextPreservingFormat(xmlDoc, valCell.tc, si!.icMemoNumber);
        } else if (/^IC\s*Memo\s*Date$/i.test(label) && nonEmpty(si?.icMemoDate)) {
          setCellTextPreservingFormat(xmlDoc, valCell.tc, si!.icMemoDate);
        } else if (/^MEC\s*Memo\s*(?:Number|No\.?)$/i.test(label) && nonEmpty(si?.mecMemoNumber)) {
          setCellTextPreservingFormat(xmlDoc, valCell.tc, si!.mecMemoNumber);
        } else if (/^MEC\s*Memo\s*Date$/i.test(label) && nonEmpty(si?.mecMemoDate)) {
          setCellTextPreservingFormat(xmlDoc, valCell.tc, si!.mecMemoDate);
        } else if (/^Board\s*Memo\s*(?:Number|No\.?)$/i.test(label) && nonEmpty(si?.boardMemoNumber)) {
          setCellTextPreservingFormat(xmlDoc, valCell.tc, si!.boardMemoNumber);
        } else if (/^Board\s*Memo\s*Date$/i.test(label) && nonEmpty(si?.boardMemoDate)) {
          setCellTextPreservingFormat(xmlDoc, valCell.tc, si!.boardMemoDate);
        } else if (/^Sanction\s*(?:Ref\.?\s*No\.?|Number)$/i.test(label) && nonEmpty(si?.sanctionRefNumber)) {
          setCellTextPreservingFormat(xmlDoc, valCell.tc, si!.sanctionRefNumber);
        } else if (/^Sanction\s*Date$/i.test(label) && nonEmpty(si?.sanctionDate)) {
          setCellTextPreservingFormat(xmlDoc, valCell.tc, si!.sanctionDate);
        } else if (/^Dated$/i.test(label) && nonEmpty(ph['#1#'])) {
          setCellTextPreservingFormat(xmlDoc, valCell.tc, ph['#1#']);
        }
      }
      continue;
    }

    // B. A/C & Branch + Subject Bar
    if (/\bA\/C\s*:/i.test(tableText) && /\bBranch\s*:/i.test(tableText)) {
      for (const row of matrix) {
        for (const c of row) {
          if (c.hasTokens) continue;
          if (/\bA\/C\s*:/i.test(c.text) && nonEmpty(ph['#2#'])) {
            const cis = nonEmpty(ph['#69#']) || '#69#';
            setCellTextPreservingFormat(xmlDoc, c.tc, `A/C: ${ph['#2#']} (CIS: ${cis})`);
          } else if (/\bBranch\s*:/i.test(c.text) && nonEmpty(ph['#3#'])) {
            setCellTextPreservingFormat(xmlDoc, c.tc, `Branch: ${ph['#3#']}`);
          } else if (/^Subject\s*:/i.test(c.text) && nonEmpty(ph['#4#'])) {
            const updated = c.text.replace(
              /(waiver of Tk\.?\s*)([^\s,]+)/i,
              `$1${ph['#4#']}`
            );
            if (updated !== c.text) {
              setCellTextPreservingFormat(xmlDoc, c.tc, updated);
            }
          }
        }
      }
      continue;
    }

    // C. Main Settlement Table (Particulars + Grand-Total + Principal)
    if (/\bPrincipal\b/i.test(tableText) && /\bGrand[-\s]*Total\b/i.test(tableText)) {
      for (let r = 0; r < matrix.length; r++) {
        const row = matrix[r];
        const joined = row.map((c) => c.text).join(' | ');

        // Header row "O/S as on ..."
        for (const c of row) {
          if (!c.hasTokens && /O\/S\s+as\s+on\b/i.test(c.text) && nonEmpty(ph['#68#'])) {
            const updated = c.text.replace(/(O\/S\s+as\s+on\s+).*/i, `$1${ph['#68#']}`);
            setCellTextPreservingFormat(xmlDoc, c.tc, updated);
          }
        }

        // First data row containing Principal (also contains Mode, Limit Original, RSD, Remarks via rowSpan)
        const princCellIdx = row.findIndex((c) => /^Principal$/i.test(c.text.trim()));
        if (princCellIdx !== -1) {
          if (princCellIdx >= 3) {
            if (!row[0].hasTokens && nonEmpty(ph['#5#'])) {
              setCellTextPreservingFormat(xmlDoc, row[0].tc, ph['#5#']);
            }
            if (!row[1].hasTokens && nonEmpty(ph['#6#'])) {
              setCellTextPreservingFormat(xmlDoc, row[1].tc, ph['#6#']);
            }
            if (!row[2].hasTokens && nonEmpty(bp?.limitRSD)) {
              setCellTextPreservingFormat(xmlDoc, row[2].tc, bp!.limitRSD);
            }
          }
          const osCell = row[princCellIdx + 1];
          if (osCell && !osCell.hasTokens && nonEmpty(ph['#7#'])) {
            setCellTextPreservingFormat(xmlDoc, osCell.tc, ph['#7#']);
          }
          const remarksCell = row[row.length - 1];
          if (
            remarksCell &&
            !remarksCell.hasTokens &&
            /Settlement\s*Deadline|Break\s*up|Security\s*Value/i.test(remarksCell.text)
          ) {
            const v = (code: string) => nonEmpty(ph[code]) || code;
            const remarks = `1)Settlement Deadline:\n${v('#28#')}\n\n2) Break up\nPrincipal: ${v('#7#')}\nCharges: ${v('#8#')}\nCOF: ${v('#31#')}\nTotal: ${v('#32#')}\n\n3) Security Value:\nMV: ${v('#33#')}\nFSV: ${v('#34#')}`;
            setCellTextPreservingFormat(xmlDoc, remarksCell.tc, remarks);
          }
          continue;
        }

        if (/^Charges\b/i.test(row[0]?.text.trim() ?? '') && !row[1]?.hasTokens && nonEmpty(ph['#8#'])) {
          setCellTextPreservingFormat(xmlDoc, row[1].tc, ph['#8#']);
          continue;
        }
        if (/^Charged Profit$/i.test(row[0]?.text.trim() ?? '') && !row[1]?.hasTokens && nonEmpty(ph['#10#'])) {
          setCellTextPreservingFormat(xmlDoc, row[1].tc, ph['#10#']);
          continue;
        }
        if (/^Compensation$/i.test(row[0]?.text.trim() ?? '') && !row[1]?.hasTokens && nonEmpty(ph['#11#'])) {
          setCellTextPreservingFormat(xmlDoc, row[1].tc, ph['#11#']);
          continue;
        }
        if (/^Uncharged Compensation$/i.test(row[0]?.text.trim() ?? '') && !row[1]?.hasTokens && nonEmpty(ph['#12#'])) {
          setCellTextPreservingFormat(xmlDoc, row[1].tc, ph['#12#']);
          continue;
        }
        if (/Grand[-\s]*Total/i.test(joined)) {
          if (row[1] && !row[1].hasTokens && nonEmpty(ph['#25#'])) {
            setCellTextPreservingFormat(xmlDoc, row[1].tc, ph['#25#']);
          }
          if (row[2] && !row[2].hasTokens && nonEmpty(ph['#4#'])) {
            setCellTextPreservingFormat(xmlDoc, row[2].tc, ph['#4#']);
          }
          if (row[3] && !row[3].hasTokens && nonEmpty(ph['#27#'])) {
            setCellTextPreservingFormat(xmlDoc, row[3].tc, ph['#27#']);
          }
        }
      }
      continue;
    }

    // D. Down Payment Note Bar
    if (/Note:\s*Information Regarding Down payment/i.test(tableText) && matrix[0]?.[0]) {
      const c = matrix[0][0];
      if (!c.hasTokens && (nonEmpty(ph['#35#']) || nonEmpty(ph['#36#']) || nonEmpty(ph['#68#']))) {
        const dp = nonEmpty(ph['#35#']) || '#35#';
        const ac = nonEmpty(ph['#36#']) || '#36#';
        const dt = nonEmpty(ph['#68#']) || '#68#';
        setCellTextPreservingFormat(
          xmlDoc,
          c.tc,
          `Note: Information Regarding Down payment (deposited Tk.${dp} Million in A/C: ${ac} on ${dt}.)`
        );
      }
      continue;
    }

    // E. Key Points Table
    if (/Cost of Fund/i.test(tableText) && /Maximum Waiver/i.test(tableText)) {
      for (const row of matrix) {
        for (let ci = 0; ci < row.length - 1; ci++) {
          const label = row[ci].text;
          const next = row[ci + 1];
          if (!next || next.hasTokens) continue;
          if (/Total\s*\[Profit/i.test(label) && nonEmpty(ph['#42#'])) {
            setCellTextPreservingFormat(xmlDoc, next.tc, `: ${ph['#42#']}`);
          } else if (/^Cost of Fund$/i.test(label) && nonEmpty(ph['#31#'])) {
            setCellTextPreservingFormat(xmlDoc, next.tc, `: ${ph['#31#']}`);
          } else if (/Maximum Waiver/i.test(label) && nonEmpty(ph['#44#'])) {
            setCellTextPreservingFormat(xmlDoc, next.tc, `: ${ph['#44#']}`);
          } else if (/^(Client|Branch)$/i.test(label) && nonEmpty(ph['#45#'])) {
            setCellTextPreservingFormat(xmlDoc, next.tc, `: ${ph['#45#']}`);
          } else if (/^(SAMD,\s*HO|HOIC|MEC|IC\s*&\s*CD)$/i.test(label) && nonEmpty(ph['#47#'])) {
            setCellTextPreservingFormat(xmlDoc, next.tc, `: ${ph['#47#']}`);
          }
        }
      }
      continue;
    }

    // F. Section 1 Nature of Business single-cell box
    if (/^(?:1\.\s*)?Nature of Business\s*:/i.test(tableText) && matrix.length === 1 && matrix[0]?.[0] && !matrix[0][0].hasTokens) {
      if (nonEmpty(ph['#54#'])) {
        setCellTextPreservingFormat(xmlDoc, matrix[0][0].tc, `1. Nature of Business: ${ph['#54#']}`);
      }
      continue;
    }

    // G. Section 2 Liability Position Table
    if (/Name of the Concern/i.test(tableText) && /Mode of Investment/i.test(tableText)) {
      for (const row of matrix) {
        const isFundedRow =
          row.length >= 8 &&
          /\bFunded\b/i.test(row[1]?.text ?? '') &&
          !/Non[-\s]*Funded/i.test(row[1]?.text ?? '');
        const nonFundedIdx = row.findIndex((c, idx) => idx <= 1 && /^Non[-\s]*Funded$/i.test(c.text.trim()));
        const isTotalRow = /^Total$/i.test(row[0]?.text.trim() ?? '') && row.length >= 7;

        if (isFundedRow) {
          if (!row[0].hasTokens && nonEmpty(ph['#2#'])) setCellTextPreservingFormat(xmlDoc, row[0].tc, ph['#2#']);
          if (!row[1].hasTokens && nonEmpty(ph['#55#'])) setCellTextPreservingFormat(xmlDoc, row[1].tc, `${ph['#55#']}\nFunded`);
          if (!row[2].hasTokens && ph['#56#'] !== undefined && (nonEmpty(ph['#56#']) || bp !== null)) {
            setCellTextPreservingFormat(xmlDoc, row[2].tc, ph['#56#'] || '');
          }
          if (!row[3].hasTokens && nonEmpty(ph['#25#'])) setCellTextPreservingFormat(xmlDoc, row[3].tc, ph['#25#']);
          if (!row[4].hasTokens && ph['#58#'] !== undefined && (nonEmpty(ph['#58#']) || bp !== null)) {
            setCellTextPreservingFormat(xmlDoc, row[4].tc, ph['#58#'] || '');
          }
          if (!row[5].hasTokens && nonEmpty(ph['#59#'])) setCellTextPreservingFormat(xmlDoc, row[5].tc, ph['#59#']);
          if (!row[6].hasTokens && nonEmpty(ph['#60#'])) setCellTextPreservingFormat(xmlDoc, row[6].tc, ph['#60#']);
          if (!row[7].hasTokens && bp !== null) {
            setCellTextPreservingFormat(xmlDoc, row[7].tc, bp.sec2Remarks || '');
          }
        } else if (nonFundedIdx !== -1 && row.length >= 7) {
          const remarksCell = row[row.length - 1];
          if (!remarksCell.hasTokens && memo.sec2NonFundedNote !== undefined && (nonEmpty(memo.sec2NonFundedNote) || bp !== null)) {
            setCellTextPreservingFormat(xmlDoc, remarksCell.tc, memo.sec2NonFundedNote || '');
          }
        } else if (isTotalRow) {
          if (!row[1].hasTokens && nonEmpty(ph['#56#'])) setCellTextPreservingFormat(xmlDoc, row[1].tc, ph['#56#']);
          if (!row[2].hasTokens && nonEmpty(ph['#25#'])) setCellTextPreservingFormat(xmlDoc, row[2].tc, ph['#25#']);
          if (!row[3].hasTokens && nonEmpty(ph['#58#'])) setCellTextPreservingFormat(xmlDoc, row[3].tc, ph['#58#']);
          if (!row[4].hasTokens && nonEmpty(ph['#59#'])) setCellTextPreservingFormat(xmlDoc, row[4].tc, ph['#59#']);
          if (!row[5].hasTokens && nonEmpty(ph['#60#'])) setCellTextPreservingFormat(xmlDoc, row[5].tc, ph['#60#']);
        }
      }
      continue;
    }

    // G2. Section 3 Table (Liability with Other Units / Banks/NBFIs)
    if (/Name of the Banks\/NBFIs/i.test(tableText) && /CL Status/i.test(tableText)) {
      if (bp !== null && matrix.length > 1) {
        const subTotalIdx = matrix.findIndex((r, idx) => idx > 0 && /sub[-\s]*total|grand\s*total/i.test(r[0]?.text ?? ''));
        const dataEndIdx = subTotalIdx !== -1 ? subTotalIdx : matrix.length;
        const dataRows = (bp.section3Rows || [])
          .filter((r) => !/sub[-\s]*total|grand\s*total/i.test(r.sl) && !/sub[-\s]*total|grand\s*total/i.test(r.nameOfConcern))
          .map((r) => [
            r.sl,
            r.nameOfConcern,
            r.bankOrNbfi,
            r.natureOfFacility,
            r.limit,
            r.outstanding,
            r.overdueEol,
            r.expiryDate,
            r.clStatus,
          ]);
        if (dataRows.length > 0 || bp.section3Rows.length === 0) {
          populateTableDataRowsPreservingFormat(xmlDoc, tbl, 1, dataEndIdx, dataRows, true, true);
        }
      }
      continue;
    }

    // G3. Section 4 Table (CIB Status)
    if (/With respective unit/i.test(tableText) && /With Other Bank/i.test(tableText)) {
      if (bp?.section4Rows && bp.section4Rows.length > 0) {
        for (const row of matrix) {
          if (row.length < 6) continue;
          const st = row[0].text.trim().toLowerCase();
          const match = bp.section4Rows.find((r) => r.status.trim().toLowerCase() === st);
          if (match) {
            if (!row[1].hasTokens) setCellTextPreservingFormat(xmlDoc, row[1].tc, match.unitBorrower);
            if (!row[2].hasTokens) setCellTextPreservingFormat(xmlDoc, row[2].tc, match.unitGuarantor);
            if (!row[3].hasTokens) setCellTextPreservingFormat(xmlDoc, row[3].tc, match.otherBorrower);
            if (!row[4].hasTokens) setCellTextPreservingFormat(xmlDoc, row[4].tc, match.otherGuarantor);
            if (!row[5].hasTokens) setCellTextPreservingFormat(xmlDoc, row[5].tc, match.remarks);
          }
        }
      }
      continue;
    }

    // G4. Section 5 Table (Legal Action Status)
    if (/Type of Case/i.test(tableText) && /Case No\./i.test(tableText)) {
      if (bp !== null && matrix.length > 1) {
        const rowsData = (bp.section5Rows || []).map((r) => [
          r.sl,
          r.typeOfCase,
          r.caseNo,
          r.filingDate,
          r.amount,
          r.previousDate,
          r.nextDate,
          r.presentStatus,
        ]);
        populateTableDataRowsPreservingFormat(xmlDoc, tbl, 1, matrix.length, rowsData, true, true);
      }
      continue;
    }

    // G5. Section 6 Table (Audit, Compliance)
    if (/Name of the Authority/i.test(tableText) && /Objection Status/i.test(tableText)) {
      if (bp?.section6Rows && bp.section6Rows.length > 0) {
        for (const row of matrix) {
          if (row.length < 6) continue;
          const auth = row[1].text.trim().toLowerCase();
          const match = bp.section6Rows.find(
            (r) =>
              r.authority.trim().toLowerCase() === auth ||
              (r.sl && row[0].text.trim().replace(/\.$/, '').toLowerCase() === r.sl.trim().replace(/\.$/, '').toLowerCase())
          );
          if (match) {
            if (!row[2].hasTokens) setCellTextPreservingFormat(xmlDoc, row[2].tc, match.lastInspectionDate);
            if (!row[3].hasTokens) setCellTextPreservingFormat(xmlDoc, row[3].tc, match.objectionStatus);
            if (!row[4].hasTokens) setCellTextPreservingFormat(xmlDoc, row[4].tc, match.objections);
            if (!row[5].hasTokens) setCellTextPreservingFormat(xmlDoc, row[5].tc, match.branchReply);
          }
        }
      }
      continue;
    }

    // G6. Section 7 Primary Security Table
    if (/Total value of primary security/i.test(tableText)) {
      const totalIdx = matrix.findIndex((r) => /Total value of primary security/i.test(r[0]?.text ?? ''));
      if (bp?.section7PrimaryRows && bp.section7PrimaryRows.length > 0 && totalIdx > 1) {
        populateTableDataRowsPreservingFormat(
          xmlDoc,
          tbl,
          1,
          totalIdx,
          bp.section7PrimaryRows.map((r) => [r.particulars, r.value]),
          true
        );
      }
      if (totalIdx !== -1 && bp !== null) {
        const refreshedTrs = directChildren(tbl, 'tr');
        const refreshedTotalTr = refreshedTrs.find((tr) =>
          /Total value of primary security/i.test(directChildren(tr, 'tc')[0] ? getCellPlainText(directChildren(tr, 'tc')[0]) : '')
        );
        if (refreshedTotalTr) {
          const totalTcs = directChildren(refreshedTotalTr, 'tc');
          if (totalTcs.length >= 2 && !cellReplacedAnyToken(replacedTokensMap, totalTcs[1])) {
            setCellTextPreservingFormat(xmlDoc, totalTcs[1], bp.section7PrimaryTotal || '');
          }
        }
      }
      continue;
    }

    // G7. Section 7 Collateral Valuation Table
    if (/Value of the mortgaged properties/i.test(tableText) || /Lowest Security value Coverage/i.test(tableText)) {
      const normSched = (s: string) => s.trim().toLowerCase().replace(/[-\s]+0*(\d+)/g, '-$1');
      if (bp?.section7ValuationRows && bp.section7ValuationRows.length > 0) {
        for (const row of matrix) {
          if (row.length === 11) {
            const sched = normSched(row[0].text);
            const match = bp.section7ValuationRows.find((v) => normSched(v.schedule) === sched);
            if (match) {
              const vals = [
                match.origBranchMv,
                match.origBranchFsv,
                match.origSurvMv,
                match.origSurvFsv,
                match.presBranchMv,
                match.presBranchFsv,
                match.presSurvMv,
                match.presSurvFsv,
                match.presAvMv,
                match.presAvFsv,
              ];
              for (let c = 0; c < 10; c++) {
                if (!row[c + 1].hasTokens) {
                  setCellTextPreservingFormat(xmlDoc, row[c + 1].tc, vals[c] ?? '');
                }
              }
            }
          }
        }
      }
      const covRow = matrix[matrix.length - 1];
      if (
        covRow?.[0] &&
        !covRow[0].hasTokens &&
        /Lowest Security value Coverage/i.test(covRow[0].text) &&
        nonEmpty(bp?.section7CoveragePct)
      ) {
        setCellTextPreservingFormat(
          xmlDoc,
          covRow[0].tc,
          `Outstanding to Lowest Security value Coverage (%): ${bp!.section7CoveragePct}`
        );
      }
      continue;
    }

    // G8. Section 9 Sequence of HO Sanctions Table
    if (/HO Sanction Reference No\./i.test(tableText)) {
      if (bp !== null && matrix.length > 1) {
        populateTableDataRowsPreservingFormat(
          xmlDoc,
          tbl,
          1,
          matrix.length,
          (bp.section9Rows || []).map((r) => [r.sl, r.referenceNo, r.date, r.limit, r.remarks]),
          true,
          true
        );
      }
      continue;
    }

    // G9. Section 10 Proprietor / Partner / Directors Table
    if (/Name of the Directors/i.test(tableText) && /Share\s*\(%\)/i.test(tableText)) {
      if (bp !== null && matrix.length > 1) {
        populateTableDataRowsPreservingFormat(
          xmlDoc,
          tbl,
          1,
          matrix.length,
          (bp.section10Directors || []).map((d) => [
            d.name,
            d.age,
            d.education,
            d.designation,
            d.sharePct,
            d.experience,
            d.netWorth,
          ]),
          true,
          true
        );
      }
      continue;
    }

    // G10. Section 10 Bank / FI / Insurance Directors Table
    if (/Name of the bank\/NBFIs/i.test(tableText) && /Number of Shares/i.test(tableText)) {
      if (bp !== null && matrix.length > 1) {
        populateTableDataRowsPreservingFormat(
          xmlDoc,
          tbl,
          1,
          matrix.length,
          (bp.section10BankDirectors || []).map((d) => [
            d.name,
            d.bankNbfi,
            d.designation,
            d.sharePct,
            d.numShares,
          ]),
          true,
          true
        );
      }
      continue;
    }

    // G11. Section 11 Financials Table
    if (/Gross Sales/i.test(tableText) && /Total Asset/i.test(tableText) && /Net Worth/i.test(tableText)) {
      if (bp !== null && matrix.length > 1) {
        const finData =
          bp.section11Rows && bp.section11Rows.length > 0
            ? bp.section11Rows.map((r) => [
                r.year,
                r.grossSales,
                r.grossProfit,
                r.netProfit,
                r.totalAsset,
                r.totalLiabilities,
                r.netWorth,
              ])
            : bp.section11FinancialsNote
              ? [['', bp.section11FinancialsNote, '', '', '', '', '']]
              : [];
        populateTableDataRowsPreservingFormat(xmlDoc, tbl, 1, matrix.length, finData, true, true);
      }
      continue;
    }

    // H. Section 16 Particulars of the facility Table
    if (/Settlement\/Exit amount/i.test(tableText) && /Down Payment/i.test(tableText)) {
      for (const row of matrix) {
        if (row.length < 2) continue;
        const label = row[row.length - 2].text.replace(/\s*:\s*$/, '').trim();
        const val = row[row.length - 1];
        if (val.hasTokens) continue;
        const dt = nonEmpty(ph['#68#']) || '#68#';
        if (/^Outstanding$/i.test(label) && nonEmpty(ph['#25#'])) {
          setCellTextPreservingFormat(xmlDoc, val.tc, `Tk. ${ph['#25#']} million as on ${dt}.`);
        } else if (/^Down Payment$/i.test(label) && nonEmpty(ph['#35#'])) {
          const ac = nonEmpty(ph['#36#']) || '#36#';
          setCellTextPreservingFormat(xmlDoc, val.tc, `Down-payment of Tk. ${ph['#35#']} Million is kept in ${ac} A/C ${ac}`);
        } else if (/^Waiver$/i.test(label) && nonEmpty(ph['#4#'])) {
          setCellTextPreservingFormat(xmlDoc, val.tc, `Tk. ${ph['#4#']} million as on ${dt}`);
        } else if (/Settlement\s*\/\s*Exit amount/i.test(label) && nonEmpty(ph['#27#'])) {
          setCellTextPreservingFormat(xmlDoc, val.tc, `Tk. ${ph['#27#']} Million as on ${dt}`);
        } else if (/^Settlement period$/i.test(label) && nonEmpty(ph['#28#'])) {
          setCellTextPreservingFormat(xmlDoc, val.tc, ph['#28#']);
        }
      }
      continue;
    }

    // I. Generic Key-Value rows (2-column [Label, Value] or 3/4-column [SL, Label, (:), Value])
    for (const row of matrix) {
      if (row.length < 2 || row.length > 4) continue;
      const valCell = row[row.length - 1];
      if (!valCell || valCell.hasTokens) continue;

      // Determine label cell (skip ":" separator cell)
      let labelCell = row[row.length - 2];
      if (labelCell.text.trim() === ':' && row.length >= 3) {
        labelCell = row[row.length - 3];
      }
      const rawLabel = labelCell.text.trim();
      if (!rawLabel) continue;
      // Normalize optional leading serial ("a. ", "1. ") and trailing ":" + collapse whitespace
      const label = rawLabel
        .replace(/^(?:[a-z]\.|\d+\.)\s+/i, '')
        .replace(/\s*:\s*$/, '')
        .replace(/\s+/g, ' ')
        .trim();
      if (!label) continue;

      const sp = bp?.section10Particulars;
      const trySet = (val: string | null) => {
        if (val !== null) setCellTextPreservingFormat(xmlDoc, valCell.tc, val);
      };

      if (/^Name of the Client$/i.test(label)) trySet(nonEmpty(sp?.nameOfClient) || clientName);
      else if (/^Unique Customer ID$/i.test(label)) trySet(nonEmpty(sp?.uniqueCustomerId) || nonEmpty(ph['#69#']));
      else if (/^AWCD\/MSD Account/i.test(label)) trySet(nonEmpty(sp?.awcdMsdAccount));
      else if (/^Business Type\/Nature$/i.test(label)) trySet(nonEmpty(sp?.businessTypeNature) || nonEmpty(ph['#54#']));
      else if (/^Business Management Type$/i.test(label)) trySet(nonEmpty(sp?.businessManagementType));
      else if (/^Client Type$/i.test(label)) trySet(nonEmpty(sp?.clientType));
      else if (/^Key Person/i.test(label)) trySet(nonEmpty(sp?.keyPersonPhone));
      else if (/^Address$/i.test(label)) trySet(nonEmpty(sp?.address));
      else if (/^Office\/Business Address$/i.test(label)) trySet(nonEmpty(sp?.officeBusinessAddress));
      else if (/^Factory Address$/i.test(label)) trySet(nonEmpty(sp?.factoryAddress));
      else if (/^Purpose of Investment$/i.test(label)) trySet(nonEmpty(sp?.purposeOfInvestment));
      else if (/^Limit of Investment$/i.test(label)) trySet(nonEmpty(sp?.limitOfInvestment) || nonEmpty(ph['#6#']));
      else if (/^Mode of Investment$/i.test(label)) trySet(nonEmpty(sp?.modeOfInvestment) || nonEmpty(ph['#5#']));
      else if (/^Last Visit Date$/i.test(label)) trySet(nonEmpty(sp?.lastVisitDate));
      else if (/^Stock Position$/i.test(label)) trySet(nonEmpty(sp?.stockPosition));
      else if (/^Fixed Asset$/i.test(label)) trySet(nonEmpty(sp?.fixedAsset));
      else if (/^Business Net Worth$/i.test(label)) trySet(nonEmpty(sp?.businessNetWorth));
      else if (/^Insurance Coverage$/i.test(label)) trySet(nonEmpty(sp?.insuranceCoverage));
      else if (/^Trade License/i.test(label)) trySet(nonEmpty(sp?.tradeLicenseValidity));
      else if (/^Other Regulatory Documents$/i.test(label)) trySet(nonEmpty(sp?.otherRegulatoryDocuments));
      else if (/^Date of Original Investment$/i.test(label)) trySet(nonEmpty(sp?.dateOfOriginalInvestment));
      else if (/Disbursement Manager$/i.test(label)) trySet(nonEmpty(sp?.firstDisbursementManager));
      else if (/First Date of Classification/i.test(label)) trySet(nonEmpty(sp?.firstDateOfClassification));
      else if (/^(Account\s*No\.?|A\/C\s*No\.?)$/i.test(label)) trySet(nonEmpty(ph['#2#']));
      else if (/^(Branch|Branch\s*Name)$/i.test(label)) trySet(nonEmpty(ph['#3#']));
      else if (/^(Proposal\s*Date|Sanction\s*Date|Memo\s*Date|Dated)$/i.test(label)) trySet(nonEmpty(ph['#1#']));
      else if (/^(Proposed\s*Waiver|Waiver\s*Amount)$/i.test(label)) trySet(nonEmpty(ph['#4#']));
      else if (/^(Settlement\s*Amount|Exit\s*Amount)$/i.test(label)) trySet(nonEmpty(ph['#27#']));
      else if (/^(Settlement\s*Period|Settlement\s*Deadline)$/i.test(label)) trySet(nonEmpty(ph['#28#']));
      else if (/^(?:8\s*[.)]\s*)?Source\s+of\s+Repayment$/i.test(label)) trySet(nonEmpty(memo.sourceOfRepayment));
      else if (/^(?:12\s*[.)]\s*)?Contents\s+of\s+(?:the\s+)?Client[’']?s\s+Prayer$/i.test(label)) trySet(nonEmpty(memo.clientPrayer));
      else if (/^(?:12\s*[.)]\s*)?Branch\s+Visit\s+Report$/i.test(label)) trySet(nonEmpty(memo.branchVisitReport));
      else if (/^(?:13\s*[.)]\s*)?Recommendation\s+of\s+(?:the\s+)?Branch$/i.test(label)) trySet(nonEmpty(memo.sec13BranchRecommendation));
      else if (/^(?:14\s*[.)]\s*)?Recommendation\s+of\s+(?:Internal\s+Control|IC\s*&?\s*CD|ICCD)/i.test(label)) {
        trySet(nonEmpty(memo.sec14IccdRecommendation));
      } else if (/^(?:15\s*[.)]\s*)?(?:Observations?\s*(?:and|&|\/)\s*Recommendations?\s*of\s*SAMD|SAMD\s*Observations?(?:\s*(?:and|&|\/)\s*Recommendations?)?)$/i.test(label)) {
        trySet(nonEmpty(memo.samdObservation));
      } else if (/^Justification\s+of\s+(?:the\s+)?Proposed\s+Facility$/i.test(label)) trySet(nonEmpty(memo.justificationLines));
      else if (stageMemo) {
        const si = stageMemo.stageInputs;
        const ss = stageMemo.stageSections;
        if (/^(?:Chronologic\s+Order\s+of\s+(?:the\s+)?IC\s*Memo|IC\s*Memo\s*(?:Number|No\.?))$/i.test(label)) trySet(nonEmpty(si.icMemoNumber));
        else if (/^IC\s*Memo\s*Date$/i.test(label)) trySet(nonEmpty(si.icMemoDate));
        else if (/^(?:Chronologic\s+Order\s+of\s+(?:the\s+)?MEC\s*Memo|MEC\s*Memo\s*(?:Number|No\.?))$/i.test(label)) trySet(nonEmpty(si.mecMemoNumber));
        else if (/^MEC\s*Memo\s*Date$/i.test(label)) trySet(nonEmpty(si.mecMemoDate));
        else if (/^(?:Chronologic\s+Order\s+of\s+(?:the\s+)?Board\s*Memo|Board\s*Memo\s*(?:Number|No\.?))$/i.test(label)) trySet(nonEmpty(si.boardMemoNumber));
        else if (/^Board\s*Memo\s*Date$/i.test(label)) trySet(nonEmpty(si.boardMemoDate));
        else if (/^Sanction\s*(?:Ref\.?\s*No\.?|Number)$/i.test(label)) trySet(nonEmpty(si.sanctionRefNumber));
        else if (/^Sanction\s*Date$/i.test(label)) trySet(nonEmpty(si.sanctionDate));
        else if (/(?:Recommendation\s+of\s+(?:the\s+)?Investment\s+Committee|Investment\s+Committee\s*(?:\(IC\))?\s+Recommendation)/i.test(label)) trySet(nonEmpty(ss.mecIcRecommendation));
        else if (/(?:Recommendation\s+of\s+(?:the\s+)?Management\s+Executive\s+Committee|Management\s+Executive\s+Committee\s*(?:\(MEC\))?\s+Recommendation)/i.test(label)) trySet(nonEmpty(ss.boardMecRecommendation));
        else if (/AI\s+Cross-Verified\s+Facility/i.test(label)) trySet(nonEmpty(ss.sanctionCrossVerifiedSummary));
      }
    }
  }

  // 3. Final safety pass: ensure any narrative heading still missing its populated value receives it cleanly
  const refreshedParas = Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'p'));
  const currentDocText = normalizeVerifyText(
    refreshedParas.map((p) => getParagraphPlainText(p).trim()).filter(Boolean).join('\n')
  );

  const ensureHeadingValue = (
    val: string,
    headingRegex: RegExp,
    preferSub: boolean
  ) => {
    const trimmed = val.trim();
    if (!trimmed) return;
    const firstLine = trimmed
      .replace(/\r\n/g, '\n')
      .split('\n')
      .map((l) => l.trim())
      .find(Boolean);
    if (!firstLine) return;
    const normLine = normalizeVerifyText(firstLine);
    const probe = normLine.length > 48 ? normLine.slice(0, 48).trim() : normLine;
    if (currentDocText.includes(probe)) return;

    const matchingParas = refreshedParas.filter((p) => {
      if (!p.parentNode) return false;
      const nt = normalizeVerifyText(getParagraphPlainText(p));
      if (!headingRegex.test(nt)) return false;
      // Skip Table-of-Contents lines ending with dots + page number
      if (/\.{3,}\s*\d+$/.test(nt)) return false;
      return true;
    });
    if (matchingParas.length === 0) return;
    const targetP =
      matchingParas.find((p) => normalizeVerifyText(getParagraphPlainText(p)).includes(':')) ??
      matchingParas[0];
    populateNarrativeHeading(xmlDoc, targetP, trimmed, preferSub);
  };

  ensureHeadingValue(
    memo.sourceOfRepayment,
    /^(?:8\s*[.)]\s*)?Source\s+of\s+Repayment\s*:?/i,
    false
  );
  ensureHeadingValue(
    memo.clientPrayer,
    /^(?:12\s*[.)]\s*)?Contents\s+of\s+(?:the\s+)?Client[’']?s\s+Prayer\s*:?/i,
    false
  );
  ensureHeadingValue(
    memo.branchVisitReport,
    /^(?:12\s*[.)]\s*)?Branch\s+Visit\s+Report\s*:?/i,
    false
  );
  ensureHeadingValue(
    memo.sec13BranchRecommendation,
    /^(?:13\s*[.)]\s*)?Recommendation\s+of\s+(?:the\s+)?Branch\b/i,
    true
  );
  ensureHeadingValue(
    memo.sec14IccdRecommendation,
    /^(?:14\s*[.)]\s*)?Recommendation\s+of\s+(?:Internal\s+Control|IC\s*&?\s*CD|ICCD)\b/i,
    false
  );
  ensureHeadingValue(
    memo.samdObservation,
    /^(?:15\s*[.)]\s*)?(?:Observations?\s*(?:and|&|\/)\s*Recommendations?\s*of\s*SAMD|SAMD\s*Observations?(?:\s*(?:and|&|\/)\s*Recommendations?)?)\b/i,
    true
  );

  if (stageMemo) {
    const ss = stageMemo.stageSections;
    ensureHeadingValue(ss.mecIcRecommendation, /(?:Recommendation\s+of\s+(?:the\s+)?Investment\s+Committee|Investment\s+Committee\s*(?:\(IC\))?\s+Recommendation)/i, true);
    ensureHeadingValue(ss.boardMecRecommendation, /(?:Recommendation\s+of\s+(?:the\s+)?Management\s+Executive\s+Committee|Management\s+Executive\s+Committee\s*(?:\(MEC\))?\s+Recommendation)/i, true);
    ensureHeadingValue(ss.sanctionCrossVerifiedSummary, /AI\s+Cross-Verified\s+Facility/i, true);
  }
}

export async function inspectDocxTemplate(
  file: File,
  docType: MemoDocType
): Promise<MemoTemplateRecord> {
  const lower = file.name.toLowerCase();
  if (!lower.endsWith('.docx')) {
    throw new Error('Please upload a Microsoft Word (.docx) file for exact template layout preservation.');
  }

  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  const zip = await JSZip.loadAsync(bytes);

  const docFile = zip.file('word/document.xml');
  if (!docFile) {
    throw new Error('Invalid .docx file: missing word/document.xml.');
  }

  const parser = new DOMParser();
  const tokenRegex = /(#\d+#|#[a-zA-Z][a-zA-Z0-9_]*#|\{\{[a-zA-Z0-9_.]+\}\}|\[[a-zA-Z][a-zA-Z0-9_]*\])/g;
  const detectedSet = new Set<string>();
  const previewSections: TemplatePreviewBlock[] = [];

  const xmlPaths = Object.keys(zip.files).filter((p) =>
    /^word\/(document|header\d*|footer\d*)\.xml$/i.test(p)
  );

  for (const xp of xmlPaths) {
    const xmlStr = await zip.files[xp].async('string');
    const xmlDoc = parser.parseFromString(xmlStr, 'application/xml');
    const pEls = Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'p'));
    for (const pEl of pEls) {
      const text = getParagraphPlainText(pEl);
      let m: RegExpExecArray | null;
      while ((m = tokenRegex.exec(text)) !== null) {
        detectedSet.add(m[1]);
      }
      tokenRegex.lastIndex = 0;
    }

    if (xp === 'word/document.xml') {
      const bodyEl = firstDescendant(xmlDoc.documentElement, 'body') ?? xmlDoc.documentElement;
      for (let i = 0; i < bodyEl.childNodes.length; i++) {
        const node = bodyEl.childNodes[i];
        if (node.nodeType !== 1) continue;
        const el = node as Element;
        if (el.localName === 'p') {
          const t = getParagraphPlainText(el).trim();
          if (t) previewSections.push({ type: 'paragraph', text: t });
        } else if (el.localName === 'tbl') {
          const rows = directChildren(el, 'tr').map((tr) =>
            directChildren(tr, 'tc').map((tc) => getCellPlainText(tc))
          );
          if (rows.length > 0) {
            previewSections.push({ type: 'table', rows });
          }
        }
      }
    }
  }

  return {
    docType,
    fileName: file.name,
    fileSize: file.size,
    uploadedAt: new Date().toISOString(),
    bytes,
    detectedTokens: Array.from(detectedSet),
    previewSections,
  };
}

/**
 * Populates an uploaded .docx template strictly preserving 100% of its XML structure,
 * margins, orientation, fonts, sizes, spacing, borders, column widths, headers, and footers.
 */
export async function populateDocxTemplate(
  templateBytes: Uint8Array,
  memo: ResolvedICMemo,
  _docType: MemoDocType,
  stageMemo?: ResolvedStageMemo | null
): Promise<Blob> {
  const zip = await JSZip.loadAsync(templateBytes);
  const { requiredTokens, optionalTokens } = buildTemplateTokenMap(memo, stageMemo);
  const allTokenMap: Record<string, string> = {
    ...optionalTokens,
    ...requiredTokens,
  };
  const allTokenKeys = Object.keys(allTokenMap);

  const parser = new DOMParser();
  const serializer = new XMLSerializer();

  const xmlPaths = Object.keys(zip.files).filter((p) =>
    /^word\/(document|header\d*|footer\d*)\.xml$/i.test(p)
  );

  for (const xp of xmlPaths) {
    const origXmlStr = await zip.files[xp].async('string');
    const xmlDeclMatch = origXmlStr.match(/^<\?xml[^?]*\?>\s*/i);
    const xmlDoc = parser.parseFromString(origXmlStr, 'application/xml');

    const replacedTokensMap = new Map<Element, Set<string>>();
    const allParagraphs = Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'p'));

    for (const pEl of allParagraphs) {
      const pText = getParagraphPlainText(pEl);
      if (hasAnyKnownToken(pText, allTokenKeys)) {
        const replaced = replaceTokensInParagraph(xmlDoc, pEl, allTokenMap);
        if (replaced.size > 0) {
          replacedTokensMap.set(pEl, replaced);
        }
      }
    }

    if (xp === 'word/document.xml') {
      applyStructuralPopulations(xmlDoc, memo, replacedTokensMap, stageMemo);
    }

    let serialized = serializer.serializeToString(xmlDoc);
    if (xmlDeclMatch && !serialized.startsWith('<?xml')) {
      serialized = xmlDeclMatch[0] + serialized;
    }
    zip.file(xp, serialized);
  }

  return await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

function normalizeVerifyText(str: string): string {
  return str.replace(/\s+/g, ' ').trim();
}

/**
 * Verifies that all populated memo content from the central state is actually written
 * into the generated .docx ZIP XML before delivering the downloadable file.
 */
export async function verifyPopulatedDocxBlob(
  docxBlob: Blob,
  memo: ResolvedICMemo,
  docType: MemoDocType,
  activeTemplate: MemoTemplateRecord | null,
  stageMemo?: ResolvedStageMemo | null
): Promise<{ verified: boolean; missingLabels: string[] }> {
  const zip = await JSZip.loadAsync(docxBlob);
  const parser = new DOMParser();

  const xmlPaths = Object.keys(zip.files).filter((p) =>
    /^word\/(document|header\d*|footer\d*)\.xml$/i.test(p)
  );
  const chunks: string[] = [];
  for (const xp of xmlPaths) {
    const xmlStr = await zip.files[xp].async('string');
    const xmlDoc = parser.parseFromString(xmlStr, 'application/xml');
    const pEls = Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'p'));
    for (const pEl of pEls) {
      const t = getParagraphPlainText(pEl).trim();
      if (t) chunks.push(t);
    }
  }
  const exportedText = normalizeVerifyText(chunks.join('\n'));

  const containsSnippet = (val: string): boolean => {
    const firstLine = val
      .split('\n')
      .map((l) => l.trim())
      .find(Boolean);
    if (!firstLine) return true;
    const normLine = normalizeVerifyText(firstLine);
    const probe = normLine.length > 48 ? normLine.slice(0, 48).trim() : normLine;
    return exportedText.includes(probe);
  };

  // For uploaded templates, inspect template paragraphs once to know which headings/tokens exist in that template
  let tplParagraphs: string[] = [];
  let templateText = '';
  if (activeTemplate) {
    const tplZip = await JSZip.loadAsync(activeTemplate.bytes);
    const tplChunks: string[] = [];
    for (const xp of Object.keys(tplZip.files).filter((p) =>
      /^word\/(document|header\d*|footer\d*)\.xml$/i.test(p)
    )) {
      const s = await tplZip.files[xp].async('string');
      const d = parser.parseFromString(s, 'application/xml');
      for (const pEl of Array.from(d.getElementsByTagNameNS(W_NS, 'p'))) {
        const t = normalizeVerifyText(getParagraphPlainText(pEl));
        if (t) tplChunks.push(t);
      }
    }
    tplParagraphs = tplChunks;
    templateText = tplChunks.join('\n');
  }

  const missingLabels: string[] = [];
  const checkBuiltIn = !activeTemplate;

  const checkNarrative = (
    label: string,
    value: string,
    headingRegex: RegExp,
    tokens: string[]
  ) => {
    const trimmed = value.trim();
    if (!trimmed) return;
    const expected =
      checkBuiltIn ||
      tplParagraphs.some((p) => headingRegex.test(p)) ||
      tokens.some((tk) => templateText.includes(tk));
    if (expected && !containsSnippet(trimmed)) {
      missingLabels.push(label);
    }
  };

  const ph = memo.placeholders;
  checkNarrative(
    'Justification of The Proposed Facility',
    memo.justificationLines,
    /^(?:\d+\.\s*)?Justification\s+of\s+(?:the\s+)?Proposed\s+Facility\b/i,
    ['#51#', '#justification#', '{{justification}}', '[justification]']
  );
  checkNarrative(
    '1. Nature of Business',
    ph['#54#'] || '',
    /^(?:1\.\s*)?Nature\s+of\s+Business\s*:?/i,
    ['#54#', '#natureOfBusiness#', '{{natureOfBusiness}}', '[natureOfBusiness]']
  );
  checkNarrative(
    '8. Source of Repayment',
    memo.sourceOfRepayment,
    /^(?:8\.\s*)?Source\s+of\s+Repayment\s*:?/i,
    ['#sourceOfRepayment#', '#sec8SourceOfRepayment#', '{{sourceOfRepayment}}', '[sourceOfRepayment]']
  );
  checkNarrative(
    "12. Contents of the Client's Prayer",
    memo.clientPrayer,
    /^(?:12\s*[.)]\s*)?Contents\s+of\s+(?:the\s+)?Client[’']?s\s+Prayer\s*:?/i,
    ['#clientPrayer#', '#contentsOfClientsPrayer#', '{{clientPrayer}}', '{{contentsOfClientsPrayer}}', '[clientPrayer]']
  );
  checkNarrative(
    '12. Branch Visit Report',
    memo.branchVisitReport,
    /^(?:12\s*[.)]\s*)?Branch\s+Visit\s+Report\s*:?/i,
    ['#branchVisit#', '#branchVisitReport#', '{{branchVisit}}', '{{branchVisitReport}}', '[branchVisit]', '[branchVisitReport]']
  );
  checkNarrative(
    '13. Recommendation of The Branch',
    memo.sec13BranchRecommendation,
    /^(?:13\s*[.)]\s*)?Recommendation\s+of\s+(?:the\s+)?Branch\b/i,
    ['#branchRecommendation#', '#sec13BranchRecommendation#', '{{branchRecommendation}}', '{{sec13BranchRecommendation}}']
  );
  checkNarrative(
    '14. Recommendation of ICCD',
    memo.sec14IccdRecommendation,
    /^(?:14\s*[.)]\s*)?Recommendation\s+of\s+(?:Internal\s+Control|IC\s*&?\s*CD|ICCD)\b/i,
    ['#iccdRecommendation#', '#sec14IccdRecommendation#', '{{iccdRecommendation}}', '{{sec14IccdRecommendation}}']
  );
  checkNarrative(
    '15. Observation and Recommendation of SAMD',
    memo.samdObservation,
    /^(?:15\s*[.)]\s*)?(?:Observations?\s*(?:and|&|\/)\s*Recommendations?\s*of\s*SAMD|SAMD\s*Observations?(?:\s*(?:and|&|\/)\s*Recommendations?)?)\b/i,
    ['#samdObservation#', '#samd#', '{{samdObservation}}']
  );

  if (docType !== 'ic' && stageMemo) {
    const si = stageMemo.stageInputs;
    const ss = stageMemo.stageSections;
    if (docType === 'mec') {
      checkNarrative(
        'Chronologic Order of the IC Memo',
        si.icMemoNumber,
        /(?:Chronologic\s+Order\s+of\s+(?:the\s+)?IC\s*Memo|IC\s*Memo\s*(?:Number|No))/i,
        ['#icMemoNumber#', '{{icMemoNumber}}', '#icMemoNo#']
      );
      checkNarrative(
        'IC Memo Date',
        si.icMemoDate,
        /IC\s*Memo\s*Date/i,
        ['#icMemoDate#', '{{icMemoDate}}']
      );
      checkNarrative(
        'Recommendation of the Investment Committee (IC)',
        ss.mecIcRecommendation,
        /(?:Recommendation\s+of\s+(?:the\s+)?Investment\s+Committee|Investment\s+Committee\s*(?:\(IC\))?\s+Recommendation)/i,
        ['#mecIcRecommendation#', '{{mecIcRecommendation}}']
      );
    } else if (docType === 'board') {
      checkNarrative(
        'Chronologic Order of the MEC Memo',
        si.mecMemoNumber,
        /(?:Chronologic\s+Order\s+of\s+(?:the\s+)?MEC\s*Memo|MEC\s*Memo\s*(?:Number|No))/i,
        ['#mecMemoNumber#', '{{mecMemoNumber}}', '#mecMemoNo#']
      );
      checkNarrative(
        'MEC Memo Date',
        si.mecMemoDate,
        /MEC\s*Memo\s*Date/i,
        ['#mecMemoDate#', '{{mecMemoDate}}']
      );
      checkNarrative(
        'Management Executive Committee (MEC) Recommendation',
        ss.boardMecRecommendation,
        /(?:Recommendation\s+of\s+(?:the\s+)?Management\s+Executive\s+Committee|Management\s+Executive\s+Committee\s*(?:\(MEC\))?\s+Recommendation)/i,
        ['#boardMecRecommendation#', '{{boardMecRecommendation}}']
      );
    } else if (docType === 'sanction') {
      checkNarrative(
        'Chronologic Order of the Board Memo',
        si.boardMemoNumber,
        /(?:Chronologic\s+Order\s+of\s+(?:the\s+)?Board\s*Memo|Board\s*Memo\s*(?:Number|No))/i,
        ['#boardMemoNumber#', '{{boardMemoNumber}}', '#boardMemoNo#']
      );
      checkNarrative(
        'Board Memo Date',
        si.boardMemoDate,
        /Board\s*Memo\s*Date/i,
        ['#boardMemoDate#', '{{boardMemoDate}}']
      );
      checkNarrative(
        'AI Cross-Verified Facility & Liability Confirmation',
        ss.sanctionCrossVerifiedSummary,
        /AI\s+Cross-Verified\s+Facility/i,
        ['#sanctionCrossVerifiedSummary#', '{{sanctionCrossVerifiedSummary}}']
      );
    }
  }

  if (checkBuiltIn) {
    for (const [code, lbl] of [
      ['#1#', 'Proposal Date'],
      ['#2#', 'Account Number / Client'],
      ['#4#', 'Waiver Amount'],
      ['#25#', 'NET Outstanding'],
      ['#27#', 'Settlement Amount'],
      ['#28#', 'Settlement Period'],
    ] as const) {
      const val = (ph[code] || '').trim();
      if (val && !containsSnippet(val)) {
        missingLabels.push(`${lbl} (${code})`);
      }
    }
  }

  return {
    verified: missingLabels.length === 0,
    missingLabels,
  };
}

function applyPreviewReplacements(text: string, tokenMap: Record<string, string>): string {
  let out = text;
  const keys = Object.keys(tokenMap).sort((a, b) => b.length - a.length);
  for (const key of keys) {
    if (out.includes(key) && tokenMap[key] !== key) {
      out = out.split(key).join(tokenMap[key]);
    }
  }
  return out;
}

/**
 * Generates live preview blocks for an uploaded .docx template using the current synchronized memo state.
 */
export function renderTemplatePreviewBlocks(
  template: MemoTemplateRecord,
  memo: ResolvedICMemo,
  stageMemo?: ResolvedStageMemo | null
): TemplatePreviewBlock[] {
  const { requiredTokens, optionalTokens } = buildTemplateTokenMap(memo, stageMemo);
  const allTokenMap = { ...optionalTokens, ...requiredTokens };
  const ph = memo.placeholders;

  return template.previewSections.map((block) => {
    if (block.type === 'paragraph') {
      let text = applyPreviewReplacements(block.text || '', allTokenMap);
      if (/^(?:1\.\s*)?Nature of Business\s*:/i.test(text) && nonEmpty(ph['#54#']) && !block.text?.includes('#54#')) {
        text = `1. Nature of Business: ${ph['#54#']}`;
      } else if (/^(?:8\.\s*)?Source of Repayment\s*:/i.test(text) && nonEmpty(memo.sourceOfRepayment)) {
        text = `8. Source of Repayment: ${memo.sourceOfRepayment}`;
      } else if (/^(?:12\.\s*)?Contents of (?:the )?Client[’']?s Prayer\s*:/i.test(text) && nonEmpty(memo.clientPrayer)) {
        text = `12. Contents of the Client’s Prayer: ${memo.clientPrayer}`;
      } else if (/^(?:12\.\s*)?Branch Visit Report\s*:/i.test(text) && nonEmpty(memo.branchVisitReport)) {
        text = `12. Branch Visit Report: ${memo.branchVisitReport}`;
      } else if (/^(?:13\.\s*)?Recommendation of (?:the )?Branch\s*:/i.test(text) && nonEmpty(memo.sec13BranchRecommendation)) {
        text = `${text}\n${memo.sec13BranchRecommendation}`;
      } else if (/^(?:14\.\s*)?Recommendation of (?:Internal Control|IC\s*&?\s*CD|ICCD)/i.test(text) && nonEmpty(memo.sec14IccdRecommendation)) {
        text = `14. Recommendation of Internal Control & Compliance Division: ${memo.sec14IccdRecommendation}`;
      } else if (/^(?:15\.\s*)?(?:Observation\s*(?:and|&)\s*Recommendation\s*of\s*SAMD|SAMD\s*Observation)/i.test(text) && nonEmpty(memo.samdObservation)) {
        text = `15. Observation and Recommendation of SAMD:\n${memo.samdObservation}`;
      }
      return { type: 'paragraph', text };
    }

    const rows = (block.rows || []).map((r) =>
      r.map((cellText) => applyPreviewReplacements(cellText, allTokenMap))
    );
    return { type: 'table', rows };
  });
}
