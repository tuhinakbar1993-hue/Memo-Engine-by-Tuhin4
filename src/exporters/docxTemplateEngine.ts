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

function getCellParagraphs(tcEl: Element): Element[] {
  const paras: Element[] = [];
  for (let i = 0; i < tcEl.childNodes.length; i++) {
    const child = tcEl.childNodes[i];
    if (child.nodeType !== 1) continue;
    const el = child as Element;
    if (el.localName === 'p') {
      paras.push(el);
    } else if (el.localName === 'sdt') {
      const sdtContent = directChildren(el, 'sdtContent')[0];
      if (sdtContent) {
        paras.push(...directChildren(sdtContent, 'p'));
      }
    }
  }
  return paras;
}

function getCellPlainText(tcEl: Element): string {
  return getCellParagraphs(tcEl)
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

function getWAttr(el: Element | null, localAttr: string): string | null {
  if (!el) return null;
  return (
    el.getAttribute(`w:${localAttr}`) ??
    el.getAttributeNS(W_NS, localAttr) ??
    el.getAttribute(localAttr) ??
    null
  );
}

interface ParsedStyleDef {
  styleId: string;
  type: string;
  name: string;
  isDefault: boolean;
  basedOn: string | null;
  rPr: Element | null;
  pPr: Element | null;
}

interface TemplateStylesInfo {
  docDefaultsRPr: Element | null;
  docDefaultsPPr: Element | null;
  defaultParagraphStyleId: string | null;
  stylesById: Map<string, ParsedStyleDef>;
}

interface TemplateStyleContext {
  styles: TemplateStylesInfo;
  narrativeBodyRPr: Element | null;
  narrativeBodyPPr: Element | null;
  narrativeBodySz: string | null;
  narrativeBodyFont: string | null;
}

const EMPTY_STYLES_INFO: TemplateStylesInfo = {
  docDefaultsRPr: null,
  docDefaultsPPr: null,
  defaultParagraphStyleId: null,
  stylesById: new Map(),
};

function parseStylesXml(stylesDoc: Document | null): TemplateStylesInfo {
  if (!stylesDoc) return EMPTY_STYLES_INFO;
  const root = stylesDoc.documentElement;
  const docDefaults = directChildren(root, 'docDefaults')[0] ?? null;
  const rPrDefault = docDefaults ? directChildren(docDefaults, 'rPrDefault')[0] : null;
  const pPrDefault = docDefaults ? directChildren(docDefaults, 'pPrDefault')[0] : null;
  const docDefaultsRPr = rPrDefault ? directChildren(rPrDefault, 'rPr')[0] ?? null : null;
  const docDefaultsPPr = pPrDefault ? directChildren(pPrDefault, 'pPr')[0] ?? null : null;

  const stylesById = new Map<string, ParsedStyleDef>();
  let defaultParagraphStyleId: string | null = null;

  for (const stEl of directChildren(root, 'style')) {
    const styleId = getWAttr(stEl, 'styleId');
    if (!styleId) continue;
    const type = getWAttr(stEl, 'type') ?? 'paragraph';
    const defAttr = getWAttr(stEl, 'default');
    const isDefault = defAttr === '1' || defAttr === 'true';
    const nameEl = directChildren(stEl, 'name')[0] ?? null;
    const name = getWAttr(nameEl, 'val') ?? styleId;
    const basedOnEl = directChildren(stEl, 'basedOn')[0] ?? null;
    const basedOn = getWAttr(basedOnEl, 'val');
    const rPr = directChildren(stEl, 'rPr')[0] ?? null;
    const pPr = directChildren(stEl, 'pPr')[0] ?? null;

    if (type === 'paragraph' && isDefault && !defaultParagraphStyleId) {
      defaultParagraphStyleId = styleId;
    }
    stylesById.set(styleId, { styleId, type, name, isDefault, basedOn, rPr, pPr });
  }

  return {
    docDefaultsRPr,
    docDefaultsPPr,
    defaultParagraphStyleId,
    stylesById,
  };
}

function isToggleOff(el: Element): boolean {
  const v = getWAttr(el, 'val');
  return v === '0' || v === 'false' || v === 'off';
}

function mergeRPrElements(
  xmlDoc: Document,
  base: Element | null,
  overlay: Element | null
): Element | null {
  if (!base && !overlay) return null;
  if (!base) return overlay!.cloneNode(true) as Element;
  if (!overlay) return base.cloneNode(true) as Element;

  const merged = xmlDoc.importNode(base, true) as Element;
  const overlayClone = xmlDoc.importNode(overlay, true) as Element;

  for (let i = 0; i < overlayClone.childNodes.length; i++) {
    const child = overlayClone.childNodes[i];
    if (child.nodeType !== 1) continue;
    const cEl = child as Element;
    const tag = cEl.localName;

    // Remove existing child of same localName before applying overlay
    for (const existing of directChildren(merged, tag)) {
      merged.removeChild(existing);
    }

    if (tag === 'rFonts') {
      // Merge font attributes from base + overlay without letting base theme attrs override explicit overlay fonts
      const baseFonts = directChildren(base, 'rFonts')[0];
      if (baseFonts) {
        const themePairs: Array<[string, string]> = [
          ['ascii', 'asciiTheme'],
          ['hAnsi', 'hAnsiTheme'],
          ['cs', 'csTheme'],
          ['eastAsia', 'eastAsiaTheme'],
        ];
        for (const [fontAttr, themeAttr] of themePairs) {
          const ovFont = getWAttr(cEl, fontAttr);
          const ovTheme = getWAttr(cEl, themeAttr);
          if (!ovFont && !ovTheme) {
            const bFont = getWAttr(baseFonts, fontAttr);
            const bTheme = getWAttr(baseFonts, themeAttr);
            if (bFont) cEl.setAttribute(`w:${fontAttr}`, bFont);
            if (bTheme) cEl.setAttribute(`w:${themeAttr}`, bTheme);
          } else if (!ovFont && ovTheme) {
            const bFont = getWAttr(baseFonts, fontAttr);
            if (bFont) cEl.setAttribute(`w:${fontAttr}`, bFont);
          }
        }
      }
    }

    // If explicitly turned off (w:val="0"/"false"/"none"), keep tag removed for clean run formatting
    if ((tag === 'b' || tag === 'bCs' || tag === 'i' || tag === 'iCs') && isToggleOff(cEl)) {
      continue;
    }
    if (tag === 'u' && getWAttr(cEl, 'val') === 'none') {
      continue;
    }
    merged.appendChild(cEl);
  }
  return merged;
}

function resolveStyleChain(
  styles: TemplateStylesInfo,
  styleId: string | null
): ParsedStyleDef[] {
  const chain: ParsedStyleDef[] = [];
  const visited = new Set<string>();
  let cur = styleId;
  while (cur && !visited.has(cur)) {
    visited.add(cur);
    const def = styles.stylesById.get(cur);
    if (!def) break;
    chain.unshift(def);
    cur = def.basedOn;
  }
  return chain;
}

function resolveStyleRPr(
  xmlDoc: Document,
  styles: TemplateStylesInfo,
  styleId: string | null,
  includeDocDefaults = true
): Element | null {
  let acc: Element | null = includeDocDefaults ? styles.docDefaultsRPr : null;
  for (const def of resolveStyleChain(styles, styleId)) {
    if (def.rPr) {
      acc = mergeRPrElements(xmlDoc, acc, def.rPr);
    }
  }
  return acc;
}

function isHeadingStyleDef(styles: TemplateStylesInfo, styleId: string | null): boolean {
  if (!styleId) return false;
  for (const def of resolveStyleChain(styles, styleId)) {
    if (/^(?:heading|title|subtitle)/i.test(def.name) || /^(?:heading|title|subtitle)/i.test(def.styleId)) {
      return true;
    }
    if (def.pPr && directChildren(def.pPr, 'outlineLvl').length > 0) {
      return true;
    }
  }
  return false;
}

function getEffectiveRunRPr(
  xmlDoc: Document,
  styles: TemplateStylesInfo,
  pEl: Element,
  rEl: Element | null
): Element | null {
  const pPr = directChildren(pEl, 'pPr')[0] ?? null;
  const pStyleEl = pPr ? directChildren(pPr, 'pStyle')[0] ?? null : null;
  const pStyleId = getWAttr(pStyleEl, 'val') ?? styles.defaultParagraphStyleId;
  let acc = resolveStyleRPr(xmlDoc, styles, pStyleId, true);

  const pRPr = pPr ? directChildren(pPr, 'rPr')[0] ?? null : null;
  if (pRPr) {
    acc = mergeRPrElements(xmlDoc, acc, pRPr);
  }

  if (rEl) {
    const rPr = directChildren(rEl, 'rPr')[0] ?? null;
    const rStyleEl = rPr ? directChildren(rPr, 'rStyle')[0] ?? null : null;
    const rStyleId = getWAttr(rStyleEl, 'val');
    if (rStyleId) {
      const charRPr = resolveStyleRPr(xmlDoc, styles, rStyleId, false);
      if (charRPr) acc = mergeRPrElements(xmlDoc, acc, charRPr);
    }
    if (rPr) {
      acc = mergeRPrElements(xmlDoc, acc, rPr);
    }
  }

  return acc;
}

function getRPrSz(rPr: Element | null): string | null {
  if (!rPr) return null;
  const szEl = directChildren(rPr, 'sz')[0] ?? null;
  const sz = getWAttr(szEl, 'val');
  if (sz) return sz;
  const szCsEl = directChildren(rPr, 'szCs')[0] ?? null;
  return getWAttr(szCsEl, 'val');
}

function getRPrFont(rPr: Element | null): string | null {
  if (!rPr) return null;
  const rfEl = directChildren(rPr, 'rFonts')[0] ?? null;
  return (
    getWAttr(rfEl, 'ascii') ??
    getWAttr(rfEl, 'hAnsi') ??
    getWAttr(rfEl, 'cs') ??
    getWAttr(rfEl, 'eastAsia') ??
    null
  );
}

function isRunBoldOrUnderlined(rPr: Element | null): boolean {
  if (!rPr) return false;
  const bEl = directChildren(rPr, 'b')[0];
  const isBold = Boolean(bEl && !isToggleOff(bEl));
  const uEl = directChildren(rPr, 'u')[0];
  const isUnd = Boolean(uEl && getWAttr(uEl, 'val') !== 'none');
  return isBold || isUnd;
}

/**
 * Strips bold/underline/italic/rStyle from an effective <w:rPr> and keeps <w:sz>/<w:szCs> synchronized
 * so populated body text never inherits heading emphasis or complex-script font-size jumps.
 */
function normalizeCleanBodyRPr(
  xmlDoc: Document,
  rPr: Element | null
): Element | null {
  if (!rPr) return null;
  const cloned = xmlDoc.importNode(rPr, true) as Element;
  for (const tag of ['rStyle', 'b', 'bCs', 'u', 'i', 'iCs']) {
    for (const el of directChildren(cloned, tag)) {
      cloned.removeChild(el);
    }
  }
  const szEl = directChildren(cloned, 'sz')[0] ?? null;
  const szCsEl = directChildren(cloned, 'szCs')[0] ?? null;
  const szVal = getWAttr(szEl, 'val') ?? getWAttr(szCsEl, 'val');
  if (szVal) {
    if (!szEl) {
      const newSz = xmlDoc.createElementNS(W_NS, 'w:sz');
      newSz.setAttribute('w:val', szVal);
      cloned.appendChild(newSz);
    }
    if (!szCsEl) {
      const newSzCs = xmlDoc.createElementNS(W_NS, 'w:szCs');
      newSzCs.setAttribute('w:val', szVal);
      cloned.appendChild(newSzCs);
    } else if (getWAttr(szCsEl, 'val') !== szVal) {
      szCsEl.setAttribute('w:val', szVal);
    }
  }
  return cloned;
}

function syncRPrSzAndSzCs(xmlDoc: Document, rPr: Element | null): Element | null {
  if (!rPr) return null;
  const cloned = xmlDoc.importNode(rPr, true) as Element;
  const szEl = directChildren(cloned, 'sz')[0] ?? null;
  const szCsEl = directChildren(cloned, 'szCs')[0] ?? null;
  const szVal = getWAttr(szEl, 'val') ?? getWAttr(szCsEl, 'val');
  if (szVal) {
    if (!szEl) {
      const newSz = xmlDoc.createElementNS(W_NS, 'w:sz');
      newSz.setAttribute('w:val', szVal);
      cloned.appendChild(newSz);
    }
    if (!szCsEl) {
      const newSzCs = xmlDoc.createElementNS(W_NS, 'w:szCs');
      newSzCs.setAttribute('w:val', szVal);
      cloned.appendChild(newSzCs);
    } else if (getWAttr(szCsEl, 'val') !== szVal) {
      szCsEl.setAttribute('w:val', szVal);
    }
  }
  return cloned;
}

function getParagraphOwnBodyRPr(
  xmlDoc: Document,
  styles: TemplateStylesInfo,
  pEl: Element
): Element | null {
  const runs = directChildren(pEl, 'r');
  for (const r of runs) {
    const rText = directChildren(r, 't')
      .map((t) => t.textContent ?? '')
      .join('')
      .trim();
    if (!rText) continue;
    const eff = getEffectiveRunRPr(xmlDoc, styles, pEl, r);
    if (!isRunBoldOrUnderlined(eff)) {
      return normalizeCleanBodyRPr(xmlDoc, eff);
    }
  }
  return null;
}

function findRepresentativeRPr(
  pEl: Element,
  templateCtx?: TemplateStyleContext | null,
  xmlDoc?: Document
): Element | null {
  const runs = directChildren(pEl, 'r');
  const textRun =
    runs.find((r) => directChildren(r, 't').some((t) => (t.textContent ?? '').trim().length > 0)) ??
    runs.find((r) => directChildren(r, 't').some((t) => (t.textContent ?? '').length > 0)) ??
    runs[0] ??
    firstDescendant(pEl, 'r');
  const pPr = directChildren(pEl, 'pPr')[0] ?? null;
  const runRPr = textRun ? firstDescendant(textRun, 'rPr') : null;
  const pRPr = pPr ? firstDescendant(pPr, 'rPr') : null;

  if (runRPr) {
    if (templateCtx && xmlDoc) {
      const eff = getEffectiveRunRPr(xmlDoc, templateCtx.styles, pEl, textRun);
      return syncRPrSzAndSzCs(xmlDoc, eff);
    }
    return runRPr.cloneNode(true) as Element;
  }

  // An empty <w:p> often inherits bold/underline or heading size inside <w:pPr><w:rPr> from pressing Enter.
  // Resolve via styles + check against templateCtx rather than leaking broken empty-paragraph styling.
  if (pRPr) {
    if (templateCtx && xmlDoc) {
      const eff = getEffectiveRunRPr(xmlDoc, templateCtx.styles, pEl, null);
      const effSz = getRPrSz(eff);
      const pTxt = getParagraphPlainText(pEl).trim();
      if (!pTxt && isRunBoldOrUnderlined(eff)) {
        if (templateCtx.narrativeBodyRPr) {
          return templateCtx.narrativeBodyRPr.cloneNode(true) as Element;
        }
        return normalizeCleanBodyRPr(xmlDoc, eff);
      }
      if (
        !pTxt &&
        pEl.parentElement?.localName !== 'tc' &&
        templateCtx.narrativeBodySz &&
        (!effSz || effSz !== templateCtx.narrativeBodySz)
      ) {
        if (templateCtx.narrativeBodyRPr) {
          return templateCtx.narrativeBodyRPr.cloneNode(true) as Element;
        }
      }
      return syncRPrSzAndSzCs(xmlDoc, eff);
    }
    return pRPr.cloneNode(true) as Element;
  }

  if (templateCtx && xmlDoc) {
    const pStyleEl = pPr ? directChildren(pPr, 'pStyle')[0] ?? null : null;
    const pStyleId = getWAttr(pStyleEl, 'val');
    if (pStyleId && !isHeadingStyleDef(templateCtx.styles, pStyleId)) {
      const stRPr = resolveStyleRPr(xmlDoc, templateCtx.styles, pStyleId, true);
      if (getRPrSz(stRPr)) {
        return syncRPrSzAndSzCs(xmlDoc, stRPr);
      }
    }
    if (templateCtx.narrativeBodyRPr) {
      return templateCtx.narrativeBodyRPr.cloneNode(true) as Element;
    }
  }

  return null;
}

/**
 * Replaces the full visible text of a paragraph while keeping <w:pPr>, bookmarks, and representative <w:rPr> intact.
 */
function setParagraphTextPreservingStyle(
  xmlDoc: Document,
  pEl: Element,
  newText: string,
  fallbackRPr?: Element | null,
  templateCtx?: TemplateStyleContext | null
) {
  const pPr = directChildren(pEl, 'pPr')[0] ?? null;
  let rPrTemplate = findRepresentativeRPr(pEl, templateCtx, xmlDoc);

  // If paragraph has no runs with text and fallbackRPr was resolved explicitly for this section,
  // apply fallbackRPr whenever the paragraph lacks an explicit clean non-bold <w:sz> (or lives outside a pre-sized cell).
  const existingNonEmptyRun = directChildren(pEl, 'r').some((r) =>
    directChildren(r, 't').some((t) => (t.textContent ?? '').trim().length > 0)
  );
  const inTableCell = pEl.parentElement?.localName === 'tc';
  if (fallbackRPr) {
    const templateHasCleanCellSz =
      inTableCell && Boolean(getRPrSz(rPrTemplate)) && !isRunBoldOrUnderlined(rPrTemplate);
    if (!rPrTemplate || (!existingNonEmptyRun && !templateHasCleanCellSz && getRPrSz(fallbackRPr))) {
      rPrTemplate = xmlDoc.importNode(fallbackRPr, true) as Element;
    } else if (!getRPrSz(rPrTemplate) && getRPrSz(fallbackRPr)) {
      rPrTemplate = mergeRPrElements(xmlDoc, fallbackRPr, rPrTemplate);
    }
  }

  if (rPrTemplate) {
    rPrTemplate = syncRPrSzAndSzCs(xmlDoc, rPrTemplate);
  }

  const preservedTags = new Set(['pPr', 'bookmarkStart', 'bookmarkEnd']);
  const toRemove: ChildNode[] = [];
  for (let i = 0; i < pEl.childNodes.length; i++) {
    const child = pEl.childNodes[i];
    if (child === pPr) continue;
    if (child.nodeType === 1 && preservedTags.has((child as Element).localName)) continue;
    toRemove.push(child);
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

  const firstBookmarkEnd = directChildren(pEl, 'bookmarkEnd')[0] ?? null;
  if (firstBookmarkEnd) {
    pEl.insertBefore(rEl, firstBookmarkEnd);
  } else {
    pEl.appendChild(rEl);
  }
}

/**
 * Ensures long or multi-line cell values wrap properly without clipping or distorting fixed table columns,
 * while strictly keeping original <w:tcPr>, <w:tcW>, borders, shading, margins, alignment, and fonts intact.
 */
function ensureCellWrapsAndFitsWithoutClipping(tcEl: Element, newText: string) {
  const isLongOrMultiLine = newText.includes('\n') || newText.length > 26;
  if (!isLongOrMultiLine) return;

  const tcPr = directChildren(tcEl, 'tcPr')[0] ?? null;
  if (tcPr) {
    const noWrap = directChildren(tcPr, 'noWrap')[0] ?? null;
    if (noWrap) {
      tcPr.removeChild(noWrap);
    }
  }

  const trEl = tcEl.parentElement?.localName === 'tr' ? tcEl.parentElement : null;
  const trPr = trEl ? directChildren(trEl, 'trPr')[0] ?? null : null;
  const trHeight = trPr ? directChildren(trPr, 'trHeight')[0] ?? null : null;
  if (trHeight) {
    const hRule = trHeight.getAttribute('w:hRule') || trHeight.getAttribute('hRule');
    if (hRule === 'exact') {
      if (trHeight.hasAttribute('w:hRule')) {
        trHeight.setAttribute('w:hRule', 'atLeast');
      } else {
        trHeight.setAttribute('hRule', 'atLeast');
      }
    }
  }
}

function isParagraphJustified(pEl: Element): boolean {
  const pPr = directChildren(pEl, 'pPr')[0];
  const jc = pPr ? directChildren(pPr, 'jc')[0] : null;
  const val = jc?.getAttribute('w:val') || jc?.getAttribute('val') || '';
  return val === 'both' || val === 'distribute';
}

/**
 * Sets text inside a <w:tc> table cell while strictly preserving <w:tcPr>, <w:pPr>, <w:rPr>,
 * borders, padding, alignment, and row structure.
 */
function setCellTextPreservingFormat(
  xmlDoc: Document,
  tcEl: Element,
  newText: string,
  fallbackRPr?: Element | null,
  templateCtx?: TemplateStyleContext | null
) {
  ensureCellWrapsAndFitsWithoutClipping(tcEl, newText);
  const paras = getCellParagraphs(tcEl);
  if (paras.length === 0) {
    const pEl = xmlDoc.createElementNS(W_NS, 'w:p');
    tcEl.appendChild(pEl);
    setParagraphTextPreservingStyle(xmlDoc, pEl, newText, fallbackRPr, templateCtx);
    return;
  }

  const firstP = paras[0];
  const normalized = newText.replace(/\r\n/g, '\n');
  const lines = normalized.split('\n');

  // If the cell paragraph uses full justification ("both"/"distribute") and we have multiple lines,
  // separate lines into cloned <w:p> elements so Word doesn't stretch words before <w:br/>.
  if (lines.length > 1 && isParagraphJustified(firstP)) {
    const firstParent = firstP.parentNode;
    setParagraphTextPreservingStyle(xmlDoc, firstP, lines[0], fallbackRPr, templateCtx);
    let anchor: Element = firstP;
    for (let i = 1; i < lines.length; i++) {
      if (i < paras.length && paras[i].parentNode === firstParent) {
        setParagraphTextPreservingStyle(xmlDoc, paras[i], lines[i], fallbackRPr, templateCtx);
        anchor = paras[i];
      } else {
        const cloned = firstP.cloneNode(true) as Element;
        setParagraphTextPreservingStyle(xmlDoc, cloned, lines[i], fallbackRPr, templateCtx);
        firstParent?.insertBefore(cloned, anchor.nextSibling);
        anchor = cloned;
      }
    }
    for (let i = lines.length; i < paras.length; i++) {
      paras[i].parentNode?.removeChild(paras[i]);
    }
    return;
  }

  // Otherwise keep single-paragraph spacing & exact <w:pPr>/<w:rPr>
  setParagraphTextPreservingStyle(xmlDoc, firstP, normalized, fallbackRPr, templateCtx);
  for (let i = 1; i < paras.length; i++) {
    paras[i].parentNode?.removeChild(paras[i]);
  }
}

/**
 * Builds a clean non-bold, non-underlined <w:rPr> inheriting font/size/color from an existing paragraph
 * and/or the template's resolved narrative body formatting.
 */
function buildCleanValueRPr(
  xmlDoc: Document,
  pEl: Element,
  templateCtx?: TemplateStyleContext | null
): Element | null {
  const existingBody = templateCtx
    ? getParagraphOwnBodyRPr(xmlDoc, templateCtx.styles, pEl)
    : null;
  if (existingBody) {
    return existingBody;
  }

  const runs = directChildren(pEl, 'r');
  const pPr = directChildren(pEl, 'pPr')[0] ?? null;
  const lastRun = runs[runs.length - 1] ?? null;
  const firstRun = runs[0] ?? null;
  const targetRun = lastRun ?? firstRun;

  let resolvedFromP: Element | null = null;
  if (templateCtx) {
    resolvedFromP = normalizeCleanBodyRPr(
      xmlDoc,
      getEffectiveRunRPr(xmlDoc, templateCtx.styles, pEl, targetRun)
    );
  } else {
    const srcRPr =
      (lastRun ? firstDescendant(lastRun, 'rPr') : null) ??
      (firstRun ? firstDescendant(firstRun, 'rPr') : null) ??
      (pPr ? firstDescendant(pPr, 'rPr') : null);
    resolvedFromP = srcRPr ? normalizeCleanBodyRPr(xmlDoc, srcRPr) : null;
  }

  if (templateCtx?.narrativeBodyRPr) {
    const pStyleEl = pPr ? directChildren(pPr, 'pStyle')[0] ?? null : null;
    const pStyleId = getWAttr(pStyleEl, 'val');
    const pSz = getRPrSz(resolvedFromP);
    const bodySz = templateCtx.narrativeBodySz;
    const headingRunIsStyled = Boolean(
      targetRun && isRunBoldOrUnderlined(firstDescendant(targetRun, 'rPr'))
    );
    // If the heading uses a heading pStyle or heading bold run with a different font size than the
    // template's body text (or omits font size), use the template's own body font size & family.
    if (
      !pSz ||
      isHeadingStyleDef(templateCtx.styles, pStyleId) ||
      (headingRunIsStyled && bodySz && pSz !== bodySz)
    ) {
      return templateCtx.narrativeBodyRPr.cloneNode(true) as Element;
    }
  }

  return resolvedFromP;
}

/**
 * Splits a single <w:r> at character offset `splitAt` inside `tNode` when a label+colon and its
 * placeholder/token were merged into the same bold/underlined heading run, placing the value
 * into a separate clean run right after the label run.
 */
function insertCleanRunAfterSplit(
  xmlDoc: Document,
  pEl: Element,
  firstNode: Element,
  prefixBeforeReplacement: string,
  suffixAfterReplacement: string,
  replacement: string,
  valueRPr: Element | null
) {
  const parentRun = firstNode.parentElement;
  if (!parentRun || parentRun.localName !== 'r') {
    setTextNodeValue(firstNode, prefixBeforeReplacement + replacement + suffixAfterReplacement);
    return;
  }

  setTextNodeValue(firstNode, prefixBeforeReplacement);

  const newRun = xmlDoc.createElementNS(W_NS, 'w:r');
  if (valueRPr) {
    newRun.appendChild(xmlDoc.importNode(valueRPr, true));
  }

  const lines = replacement.replace(/\r\n/g, '\n').split('\n');
  for (let i = 0; i < lines.length; i++) {
    if (i > 0) {
      newRun.appendChild(xmlDoc.createElementNS(W_NS, 'w:br'));
    }
    const tEl = xmlDoc.createElementNS(W_NS, 'w:t');
    const lineText = i === lines.length - 1 ? lines[i] + suffixAfterReplacement : lines[i];
    setTextNodeValue(tEl, lineText);
    newRun.appendChild(tEl);
  }

  if (parentRun.parentNode === pEl) {
    pEl.insertBefore(newRun, parentRun.nextSibling);
  } else {
    parentRun.parentNode?.insertBefore(newRun, parentRun.nextSibling);
  }
}

/**
 * Replaces "Label: <old>" inside a paragraph while preserving the bold/underlined label runs,
 * and handles both existing tails AND empty tails cleanly (e.g. "12. Branch Visit Report:").
 */
function setParagraphValueAfterColon(
  xmlDoc: Document,
  pEl: Element,
  newValue: string,
  fallbackBodyRPr?: Element | null,
  templateCtx?: TemplateStyleContext | null
): boolean {
  const full = collectDirectTextNodes(pEl)
    .map((t) => t.textContent ?? '')
    .join('');
  const colonIdx = full.indexOf(':');
  if (colonIdx === -1) return false;

  const afterIdx = colonIdx + 1;
  const tail = full.slice(afterIdx);
  const resolvedValueRPr =
    fallbackBodyRPr ?? buildCleanValueRPr(xmlDoc, pEl, templateCtx);

  if (tail.trim().length > 0) {
    const leadingSpaces = tail.match(/^\s*/)?.[0] ?? ' ';
    const replaceStart = afterIdx + leadingSpaces.length;
    const isBlankPunctPlaceholder = /^[._…\-\s\[\]<>]+$/.test(tail.trim());

    // Check whether the first overlapping <w:t> also contains the heading label (i.e. colon is in same run)
    const tNodes = collectDirectTextNodes(pEl);
    let cursor = 0;
    let firstOverlap: { node: Element; start: number; end: number; text: string } | null = null;
    const overlappingAfter: Element[] = [];
    for (const node of tNodes) {
      const txt = node.textContent ?? '';
      const s = cursor;
      const e = cursor + txt.length;
      cursor = e;
      if (e > replaceStart && s < full.length) {
        if (!firstOverlap) {
          firstOverlap = { node, start: s, end: e, text: txt };
        } else {
          overlappingAfter.push(node);
        }
      }
    }

    if (firstOverlap) {
      const firstRun = firstOverlap.node.parentElement;
      const firstRunRPr = firstRun?.localName === 'r' ? firstDescendant(firstRun, 'rPr') : null;
      const sharesLabelWithHeading =
        firstOverlap.start <= colonIdx && isRunBoldOrUnderlined(firstRunRPr);
      const isHeadingStyledPlaceholder =
        isBlankPunctPlaceholder && isRunBoldOrUnderlined(firstRunRPr);

      if (sharesLabelWithHeading || isHeadingStyledPlaceholder) {
        const prefix = firstOverlap.text.slice(0, Math.max(0, replaceStart - firstOverlap.start));
        for (const extra of overlappingAfter) {
          setTextNodeValue(extra, '');
        }
        insertCleanRunAfterSplit(
          xmlDoc,
          pEl,
          firstOverlap.node,
          prefix,
          '',
          newValue,
          resolvedValueRPr
        );
        return true;
      }
    }

    replaceSpanInParagraph(xmlDoc, pEl, replaceStart, full.length, newValue);
    // If replacing an underlined blank line like "Date: ___________", remove underline from the inserted run only
    if (isBlankPunctPlaceholder) {
      const lastRuns = directChildren(pEl, 'r');
      const lastRun = lastRuns[lastRuns.length - 1];
      const rPr = lastRun ? firstDescendant(lastRun, 'rPr') : null;
      const uEl = rPr ? firstDescendant(rPr, 'u') : null;
      if (uEl && getParagraphPlainText(pEl).endsWith(newValue.split('\n')[0])) {
        uEl.parentNode?.removeChild(uEl);
      }
    }
    return true;
  }

  // Empty tail after colon: append a clean non-bold/non-underlined run preserving template body font & size
  const rEl = xmlDoc.createElementNS(W_NS, 'w:r');
  if (resolvedValueRPr) {
    rEl.appendChild(xmlDoc.importNode(resolvedValueRPr, true));
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
  const stageDate =
    stageMemo?.docType === 'mec'
      ? si?.mecMemoDate || si?.icMemoDate || ph['#1#'] || ''
      : stageMemo?.docType === 'board'
        ? si?.boardMemoDate || si?.mecMemoDate || ph['#1#'] || ''
        : stageMemo?.docType === 'sanction'
          ? si?.sanctionDate || si?.boardMemoDate || ph['#1#'] || ''
          : ph['#1#'] || '';

  const semanticMap: Record<string, string> = {
    proposalDate: ph['#1#'] || '',
    dated: stageDate,
    memoDate: stageDate,
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
  tokenMap: Record<string, string>,
  templateCtx?: TemplateStyleContext | null
): Set<string> {
  const replacedKeys = new Set<string>();
  const keys = Object.keys(tokenMap).sort((a, b) => b.length - a.length);
  if (keys.length === 0) return replacedKeys;

  const NARRATIVE_BODY_TOKEN_KEYS = new Set([
    '#sourceOfRepayment#',
    '#sec8SourceOfRepayment#',
    '{{sourceOfRepayment}}',
    '[sourceOfRepayment]',
    '#clientPrayer#',
    '#contentsOfClientsPrayer#',
    '{{clientPrayer}}',
    '{{contentsOfClientsPrayer}}',
    '[clientPrayer]',
    '#branchVisit#',
    '#branchVisitReport#',
    '{{branchVisit}}',
    '{{branchVisitReport}}',
    '[branchVisit]',
    '[branchVisitReport]',
    '#branchRecommendation#',
    '#sec13BranchRecommendation#',
    '{{branchRecommendation}}',
    '{{sec13BranchRecommendation}}',
    '#iccdRecommendation#',
    '#sec14IccdRecommendation#',
    '{{iccdRecommendation}}',
    '{{sec14IccdRecommendation}}',
    '#samdObservation#',
    '#samd#',
    '{{samdObservation}}',
    '#mecIcRecommendation#',
    '{{mecIcRecommendation}}',
    '[mecIcRecommendation]',
    '#boardMecRecommendation#',
    '{{boardMecRecommendation}}',
    '[boardMecRecommendation]',
    '#sanctionCrossVerifiedSummary#',
    '{{sanctionCrossVerifiedSummary}}',
    '[sanctionCrossVerifiedSummary]',
  ]);

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

    // If a narrative body token was placed inside the same bold/underlined heading run after ":",
    // split it into a clean body run using the template's resolved body formatting.
    if (NARRATIVE_BODY_TOKEN_KEYS.has(bestKey) && bestIdx > 0) {
      const before = fullText.slice(0, bestIdx);
      if (/:\s*$/.test(before)) {
        const tNodes = collectDirectTextNodes(pEl);
        let cursor = 0;
        let firstOverlap: { node: Element; start: number; end: number; text: string } | null = null;
        const extraOverlaps: Element[] = [];
        const endIdx = bestIdx + bestKey.length;
        for (const node of tNodes) {
          const txt = node.textContent ?? '';
          const s = cursor;
          const e = cursor + txt.length;
          cursor = e;
          if (e > bestIdx && s < endIdx) {
            if (!firstOverlap) {
              firstOverlap = { node, start: s, end: e, text: txt };
            } else {
              extraOverlaps.push(node);
            }
          }
        }
        if (firstOverlap) {
          const parentRun = firstOverlap.node.parentElement;
          const runRPr = parentRun?.localName === 'r' ? firstDescendant(parentRun, 'rPr') : null;
          if (firstOverlap.text.slice(0, Math.max(0, bestIdx - firstOverlap.start)).includes(':') && isRunBoldOrUnderlined(runRPr)) {
            const prefix = firstOverlap.text.slice(0, Math.max(0, bestIdx - firstOverlap.start));
            const lastNodeText = extraOverlaps.length > 0 ? extraOverlaps[extraOverlaps.length - 1].textContent ?? '' : firstOverlap.text;
            const lastStart = extraOverlaps.length > 0 ? cursor - lastNodeText.length : firstOverlap.start;
            const suffix = lastNodeText.slice(Math.max(0, endIdx - lastStart));
            for (const extra of extraOverlaps) setTextNodeValue(extra, '');
            const cleanBodyRPr = buildCleanValueRPr(xmlDoc, pEl, templateCtx);
            insertCleanRunAfterSplit(
              xmlDoc,
              pEl,
              firstOverlap.node,
              prefix,
              suffix,
              tokenMap[bestKey],
              cleanBodyRPr
            );
            replacedKeys.add(bestKey);
            continue;
          }
        }
      }
    }

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
  return getCellParagraphs(tcEl).some((p) => {
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
  if (/^(?:1[0-9]|[1-9])\s*[.)]\s*[A-Za-z]/.test(t)) return true;
  if (
    /^(?:Nature\s+of\s+Business|Liability\s+Position\s+of|Liability\s+of\s+the\s+Client|CIB\s+Status|Legal\s+Action\s+Status|Audit,\s*Compliance|Security\s*:|Source\s+of\s+Repayment|Sequence\s+of\s+Head\s+Office|Brief\s+Particulars\s+of\s+The\s+Client|Financials\s+of\s+The\s+Client|Contents\s+of\s+(?:the\s+)?Client[’']?s\s+Prayer|Branch\s+Visit\s+Report|Recommendation\s+of\s+(?:the\s+)?(?:Investment\s+Committee|Management\s+Executive\s+Committee|Branch)|(?:Investment\s+Committee|Management\s+Executive\s+Committee)\s*(?:\((?:IC|MEC)\))?\s+Recommendation|AI\s+Cross-Verified\s+Facility|Branch\s+Recommendation|Recommendation\s+of\s+(?:Internal\s+Control|IC\s*&?\s*CD|ICCD)|Observations?\s*(?:and|&|\/)\s*Recommendations?\s*of\s*SAMD|SAMD\s*Observations?|Particulars\s+of\s+(?:the\s+)?(?:Proposed\s+)?Facility|Special\s+Term\(s\))/i.test(
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
    /^(?:The\s+proposal\s+is\s+hereby\s+placed|Placed\s+before\s+the|Submitted\s+(?:before|for\s+kind)|Issued\s+(?:under|in\s+strict)|Conveyed\s+for\s+immediate|Union\s+Bank\s+PLC\s+Unit|Special\s+Asset\s+Management\s+Division|\([A-Z][a-z]+\s+[A-Z]|Deputy\s+Managing\s+Director|Additional\s+Managing\s+Director|Managing\s+Director\b)/i.test(
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
 * Inspects the uploaded template's own narrative body runs, sub-paragraphs, and styles.xml
 * to derive its exact native font family, font size, and paragraph formatting without arbitrary defaults.
 */
function buildTemplateStyleContext(
  xmlDoc: Document,
  styles: TemplateStylesInfo
): TemplateStyleContext {
  interface Candidate {
    weight: number;
    rPr: Element;
    pPr: Element | null;
    sz: string | null;
    font: string | null;
  }
  const candidates: Candidate[] = [];
  const allParas = Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'p'));

  const addCandidate = (pEl: Element, rPr: Element | null, weight: number) => {
    if (!rPr) return;
    const cleaned = normalizeCleanBodyRPr(xmlDoc, rPr);
    if (!cleaned) return;
    const sz = getRPrSz(cleaned);
    const font = getRPrFont(cleaned);
    if (!sz && !font) return;
    const pPr = directChildren(pEl, 'pPr')[0] ?? null;
    candidates.push({ weight, rPr: cleaned, pPr, sz, font });
  };

  for (const pEl of allParas) {
    if (pEl.parentElement?.localName === 'tc') continue;
    const txt = getParagraphPlainText(pEl).replace(/\s+/g, ' ').trim();
    if (!txt || /\.{3,}\s*\d+$/.test(txt)) continue;

    // Highest signal: narrative sub-paragraphs or non-bold body runs on narrative sections (IC/MEC/Board/Sanction)
    if (
      /^(?:\d+\s*[.)]\s*)?(?:Source\s+of\s+Repayment|Contents\s+of\s+(?:the\s+)?Client|Branch\s+Visit\s+Report|Recommendation\s+of\s+(?:the\s+)?(?:Branch|Internal\s+Control|IC\s*&?\s*CD|ICCD|Investment\s+Committee|Management\s+Executive\s+Committee)|Branch\s+Recommendation|(?:Investment\s+Committee|Management\s+Executive\s+Committee)\s*(?:\((?:IC|MEC)\))?\s+Recommendation|Observations?\s*(?:and|&|\/)\s*Recommendations?\s*of\s*SAMD|SAMD\s*Observations?)/i.test(
        txt
      )
    ) {
      const ownBody = getParagraphOwnBodyRPr(xmlDoc, styles, pEl);
      if (ownBody) addCandidate(pEl, ownBody, 8);
      for (const sp of collectContiguousSubParagraphs(pEl)) {
        const spTxt = getParagraphPlainText(sp).trim();
        if (!spTxt || isBlankOrPlaceholderText(spTxt)) continue;
        const spBody = getParagraphOwnBodyRPr(xmlDoc, styles, sp);
        if (spBody) addCandidate(sp, spBody, 10);
      }
      continue;
    }

    // Secondary strong signal: Justification bullets b./c. and Special Terms i..ix
    if (
      /^(?:[bc]\.\s*(?:Investment\s+was\s+not|The\s+client\s+is\s+not)|(?:i|ii|iii|iv|v|vi|vii|ix)\.\s+|This\s+Memo\s+is\s+placed\s+in\s+conformity|If\s+the\s+proposal\s+approved)/i.test(
        txt
      )
    ) {
      const ownBody = getParagraphOwnBodyRPr(xmlDoc, styles, pEl);
      if (ownBody) addCandidate(pEl, ownBody, 5);
    }
  }

  let bestRPr: Element | null = null;
  let bestPPr: Element | null = null;
  if (candidates.length > 0) {
    const scoreByKey = new Map<string, { score: number; sample: Candidate }>();
    for (const c of candidates) {
      const key = `${c.font ?? ''}|${c.sz ?? ''}`;
      const existing = scoreByKey.get(key);
      if (existing) {
        existing.score += c.weight;
        if (c.weight > existing.sample.weight) {
          existing.sample = c;
        }
      } else {
        scoreByKey.set(key, { score: c.weight, sample: c });
      }
    }
    let topScore = -1;
    for (const entry of scoreByKey.values()) {
      if (entry.score > topScore) {
        topScore = entry.score;
        bestRPr = entry.sample.rPr;
        bestPPr = entry.sample.pPr;
      }
    }
  }

  // Fallback to template's own Normal / default paragraph style in styles.xml if no narrative body runs existed
  if (!bestRPr) {
    const bodyStyleId =
      (styles.stylesById.has('BodyText') ? 'BodyText' : null) ??
      styles.defaultParagraphStyleId ??
      (styles.stylesById.has('Normal') ? 'Normal' : null);
    const stRPr = resolveStyleRPr(xmlDoc, styles, bodyStyleId, true);
    bestRPr = normalizeCleanBodyRPr(xmlDoc, stRPr);
  }

  return {
    styles,
    narrativeBodyRPr: bestRPr,
    narrativeBodyPPr: bestPPr,
    narrativeBodySz: getRPrSz(bestRPr),
    narrativeBodyFont: getRPrFont(bestRPr),
  };
}

function getPreferredStyledSubParagraph(
  existingSub: Element[],
  templateCtx?: TemplateStyleContext | null
): Element | null {
  // 1. Prefer a non-empty, non-placeholder existing sub-paragraph
  const realSample = existingSub.find((p) => {
    const t = getParagraphPlainText(p).trim();
    return t !== '' && !isBlankOrPlaceholderText(t);
  });
  if (realSample) return realSample;

  // 2. Otherwise any non-empty existing sub-paragraph
  const anyNonEmpty = existingSub.find((p) => getParagraphPlainText(p).trim() !== '');
  if (anyNonEmpty) return anyNonEmpty;

  // 3. Otherwise an empty placeholder sub-paragraph that already carries valid non-heading body styling
  if (templateCtx) {
    const styledBlank = existingSub.find((p) => {
      const pPr = directChildren(p, 'pPr')[0] ?? null;
      const pStyleId = getWAttr(pPr ? directChildren(pPr, 'pStyle')[0] ?? null : null, 'val');
      if (isHeadingStyleDef(templateCtx.styles, pStyleId)) return false;
      const eff = findRepresentativeRPr(p, templateCtx, p.ownerDocument ?? undefined);
      const effSz = getRPrSz(eff);
      return Boolean(eff && effSz && (!templateCtx.narrativeBodySz || effSz === templateCtx.narrativeBodySz));
    });
    if (styledBlank) return styledBlank;
  }

  return existingSub[0] ?? null;
}

function resolveSectionBodyRPr(
  xmlDoc: Document,
  templateCtx: TemplateStyleContext | null | undefined,
  headingP: Element,
  existingSub: Element[]
): Element | null {
  if (!templateCtx) {
    const realSub = existingSub.find((p) => getParagraphPlainText(p).trim() !== '');
    if (realSub) {
      const rPr = findRepresentativeRPr(realSub);
      if (rPr) return normalizeCleanBodyRPr(xmlDoc, rPr);
    }
    return buildCleanValueRPr(xmlDoc, headingP, null);
  }

  // Priority 1: Non-empty existing sub-paragraph directly under this heading
  const realSub = existingSub.find((p) => {
    const t = getParagraphPlainText(p).trim();
    return t !== '' && !isBlankOrPlaceholderText(t);
  }) ?? existingSub.find((p) => getParagraphPlainText(p).trim() !== '');

  if (realSub) {
    const subBody = getParagraphOwnBodyRPr(xmlDoc, templateCtx.styles, realSub);
    if (subBody) return subBody;
    const subRep = findRepresentativeRPr(realSub, templateCtx, xmlDoc);
    if (subRep) return normalizeCleanBodyRPr(xmlDoc, subRep);
  }

  // Priority 2: Non-bold body run already present on the heading line itself (e.g. Section 13)
  const headingBody = getParagraphOwnBodyRPr(xmlDoc, templateCtx.styles, headingP);
  if (headingBody) return headingBody;

  // Priority 3: Blank placeholder sub-paragraph if it has valid non-heading body <w:sz> matching template body
  for (const sp of existingSub) {
    const pPr = directChildren(sp, 'pPr')[0] ?? null;
    const pStyleEl = pPr ? directChildren(pPr, 'pStyle')[0] ?? null : null;
    const pStyleId = getWAttr(pStyleEl, 'val');
    if (isHeadingStyleDef(templateCtx.styles, pStyleId)) continue;
    const eff = getEffectiveRunRPr(xmlDoc, templateCtx.styles, sp, directChildren(sp, 'r')[0] ?? null);
    const effSz = getRPrSz(eff);
    const isEmptyBlankWithHeadingCarryover =
      getParagraphPlainText(sp).trim() === '' && isRunBoldOrUnderlined(eff);
    if (
      !isEmptyBlankWithHeadingCarryover &&
      effSz &&
      (!templateCtx.narrativeBodySz || effSz === templateCtx.narrativeBodySz)
    ) {
      return normalizeCleanBodyRPr(xmlDoc, eff);
    }
  }

  // Priority 4: Template's own resolved narrative body formatting (from sibling sections & styles.xml)
  if (templateCtx.narrativeBodyRPr) {
    return templateCtx.narrativeBodyRPr.cloneNode(true) as Element;
  }

  // Priority 5: Heading's cleaned <w:rPr>
  return buildCleanValueRPr(xmlDoc, headingP, templateCtx);
}

/**
 * Inserts/updates body sub-paragraphs right under `headingP`, cloning existing sub-paragraph styling
 * when available or building clean non-bold/non-underlined runs inheriting the template's exact body font/size.
 */
function populateSubParagraphLines(
  xmlDoc: Document,
  headingP: Element,
  existingSub: Element[],
  lines: string[],
  defaultIndent = '180',
  templateCtx?: TemplateStyleContext | null,
  explicitBodyRPr?: Element | null
) {
  if (lines.length === 0) return;

  // Preserve intentional trailing blank spacer paragraph if template had non-empty content followed by blank spacer(s)
  let workingSub = existingSub;
  if (existingSub.length >= 2) {
    const hasNonEmpty = existingSub.some((p) => getParagraphPlainText(p).trim() !== '');
    if (hasNonEmpty) {
      let lastNonEmpty = existingSub.length - 1;
      while (lastNonEmpty >= 0 && getParagraphPlainText(existingSub[lastNonEmpty]).trim() === '') {
        lastNonEmpty--;
      }
      if (lastNonEmpty >= 0 && lastNonEmpty < existingSub.length - 1) {
        workingSub = existingSub.slice(0, lastNonEmpty + 1);
      }
    }
  }

  // If workingSub starts with leading empty spacer(s) followed by an actual styled sample paragraph,
  // avoid letting an unstyled/tiny leading blank spacer corrupt line 0's font size or spacing.
  const firstNonEmptyIdx = workingSub.findIndex((p) => getParagraphPlainText(p).trim() !== '');
  if (firstNonEmptyIdx > 0) {
    const leadingBlanks = workingSub.slice(0, firstNonEmptyIdx);
    workingSub = workingSub.slice(firstNonEmptyIdx);
    for (const lb of leadingBlanks) {
      lb.parentNode?.removeChild(lb);
    }
  }

  const styledSub = getPreferredStyledSubParagraph(workingSub, templateCtx);
  const resolvedBodyRPr =
    explicitBodyRPr ??
    resolveSectionBodyRPr(xmlDoc, templateCtx, headingP, existingSub);

  const styledSubHasCleanPPr = Boolean(
    styledSub &&
      (() => {
        const pPr = directChildren(styledSub, 'pPr')[0] ?? null;
        if (!pPr) return false;
        const pStyleEl = directChildren(pPr, 'pStyle')[0] ?? null;
        const pStyleId = getWAttr(pStyleEl, 'val');
        if (templateCtx && isHeadingStyleDef(templateCtx.styles, pStyleId)) return false;
        return getParagraphPlainText(styledSub).trim() !== '' || pPr.childNodes.length > 0;
      })()
  );

  const buildBodyPPr = (): Element => {
    if (templateCtx?.narrativeBodyPPr) {
      const cloned = xmlDoc.importNode(templateCtx.narrativeBodyPPr, true) as Element;
      for (const oldRPr of directChildren(cloned, 'rPr')) {
        cloned.removeChild(oldRPr);
      }
      if (defaultIndent && directChildren(cloned, 'ind').length === 0) {
        const ind = xmlDoc.createElementNS(W_NS, 'w:ind');
        ind.setAttribute('w:left', defaultIndent);
        cloned.appendChild(ind);
      }
      return cloned;
    }

    const pPr = xmlDoc.createElementNS(W_NS, 'w:pPr');
    const headingPPr = directChildren(headingP, 'pPr')[0] ?? null;
    const headingJc = headingPPr ? directChildren(headingPPr, 'jc')[0] : null;
    if (headingJc) {
      pPr.appendChild(headingJc.cloneNode(true));
    }
    if (defaultIndent) {
      const ind = xmlDoc.createElementNS(W_NS, 'w:ind');
      ind.setAttribute('w:left', defaultIndent);
      pPr.appendChild(ind);
    }
    const headingSp = headingPPr ? directChildren(headingPPr, 'spacing')[0] : null;
    if (headingSp) {
      pPr.appendChild(headingSp.cloneNode(true));
    } else {
      const sp = xmlDoc.createElementNS(W_NS, 'w:spacing');
      sp.setAttribute('w:before', '15');
      sp.setAttribute('w:after', '15');
      pPr.appendChild(sp);
    }
    return pPr;
  };

  const sanitizeSubPPrRPr = (pPrEl: Element) => {
    for (const rPrInP of directChildren(pPrEl, 'rPr')) {
      pPrEl.removeChild(rPrInP);
    }
    if (resolvedBodyRPr) {
      const sectPrEl = directChildren(pPrEl, 'sectPr')[0] ?? null;
      const clonedRPr = xmlDoc.importNode(resolvedBodyRPr, true);
      if (sectPrEl) {
        pPrEl.insertBefore(clonedRPr, sectPrEl);
      } else {
        pPrEl.appendChild(clonedRPr);
      }
    }
  };

  const applyConsistentPPrToExistingSlot = (targetP: Element) => {
    if (styledSub && styledSub !== targetP && styledSubHasCleanPPr) {
      const styledPPr = directChildren(styledSub, 'pPr')[0] ?? null;
      const oldPPr = directChildren(targetP, 'pPr')[0] ?? null;
      if (styledPPr) {
        const clonedPPr = styledPPr.cloneNode(true) as Element;
        sanitizeSubPPrRPr(clonedPPr);
        if (oldPPr) {
          targetP.replaceChild(clonedPPr, oldPPr);
        } else {
          targetP.insertBefore(clonedPPr, targetP.firstChild);
        }
      }
      return;
    }
    // If targetP had a heading pStyle or heading mark rPr inside pPr, replace/clean it
    const curPPr = directChildren(targetP, 'pPr')[0] ?? null;
    if (curPPr) {
      const pStyleEl = directChildren(curPPr, 'pStyle')[0] ?? null;
      const pStyleId = getWAttr(pStyleEl, 'val');
      if (templateCtx && isHeadingStyleDef(templateCtx.styles, pStyleId) && pStyleEl) {
        curPPr.removeChild(pStyleEl);
      }
      sanitizeSubPPrRPr(curPPr);
    } else if (templateCtx?.narrativeBodyPPr) {
      targetP.insertBefore(buildBodyPPr(), targetP.firstChild);
    }
  };

  const writeCleanBodyLineIntoSlot = (targetP: Element, lineText: string) => {
    if (resolvedBodyRPr) {
      // Strip prior runs so resolvedBodyRPr governs every line uniformly without inheriting a bold lead-in span
      for (const oldR of directChildren(targetP, 'r')) {
        targetP.removeChild(oldR);
      }
    }
    setParagraphTextPreservingStyle(xmlDoc, targetP, lineText, resolvedBodyRPr, templateCtx);
  };

  const createStyledPara = (lineText: string): Element => {
    if (styledSub && styledSubHasCleanPPr) {
      const cloned = styledSub.cloneNode(true) as Element;
      const clonedPPr = directChildren(cloned, 'pPr')[0] ?? null;
      if (clonedPPr) {
        sanitizeSubPPrRPr(clonedPPr);
      }
      writeCleanBodyLineIntoSlot(cloned, lineText);
      return cloned;
    }
    const pNew = xmlDoc.createElementNS(W_NS, 'w:p');
    const pPr = buildBodyPPr();
    pNew.appendChild(pPr);

    const rEl = xmlDoc.createElementNS(W_NS, 'w:r');
    if (resolvedBodyRPr) {
      rEl.appendChild(xmlDoc.importNode(resolvedBodyRPr, true));
    }
    const tEl = xmlDoc.createElementNS(W_NS, 'w:t');
    setTextNodeValue(tEl, lineText);
    rEl.appendChild(tEl);
    pNew.appendChild(rEl);
    return pNew;
  };

  if (workingSub.length > 0) {
    applyConsistentPPrToExistingSlot(workingSub[0]);
    writeCleanBodyLineIntoSlot(workingSub[0], lines[0]);
    let insertAfter: Element = workingSub[0];
    for (let k = 1; k < lines.length; k++) {
      if (k < workingSub.length) {
        applyConsistentPPrToExistingSlot(workingSub[k]);
        writeCleanBodyLineIntoSlot(workingSub[k], lines[k]);
        insertAfter = workingSub[k];
      } else {
        const cloned = createStyledPara(lines[k]);
        insertAfter.parentNode?.insertBefore(cloned, insertAfter.nextSibling);
        insertAfter = cloned;
      }
    }
    for (let k = lines.length; k < workingSub.length; k++) {
      workingSub[k].parentNode?.removeChild(workingSub[k]);
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
  preferSubParagraphs = false,
  templateCtx?: TemplateStyleContext | null
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
  const resolvedBodyRPr = resolveSectionBodyRPr(xmlDoc, templateCtx, pEl, subParas);

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
        setCellTextPreservingFormat(xmlDoc, nextCell, lines.join('\n'), resolvedBodyRPr, templateCtx);
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
            setCellTextPreservingFormat(xmlDoc, nextRowCells[0], lines.join('\n'), resolvedBodyRPr, templateCtx);
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
      if (setParagraphValueAfterColon(xmlDoc, pEl, lines.join('\n'), resolvedBodyRPr, templateCtx)) {
        // Preserve trailing blank spacer paragraph if template had one, remove stale non-empty lines
        const nonEmptyCount = nonEmptySubParas.length;
        for (let k = 0; k < subParas.length; k++) {
          if (k === subParas.length - 1 && nonEmptyCount > 0 && getParagraphPlainText(subParas[k]).trim() === '') {
            continue;
          }
          subParas[k].parentNode?.removeChild(subParas[k]);
        }
        return;
      }
    }
    populateSubParagraphLines(xmlDoc, pEl, subParas, lines, '120', templateCtx, resolvedBodyRPr);
    return;
  }

  if (colonIdx !== -1 && !preferSubParagraphs) {
    if (setParagraphValueAfterColon(xmlDoc, pEl, lines.join('\n'), resolvedBodyRPr, templateCtx)) {
      return;
    }
  }

  populateSubParagraphLines(xmlDoc, pEl, subParas, lines, '120', templateCtx, resolvedBodyRPr);
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

interface GridCellInfo {
  tc: Element;
  text: string;
  hasTokens: boolean;
  colStart: number;
  colEnd: number;
  gridSpan: number;
  vMerge: 'restart' | 'continue' | 'none';
}

function getRowGridCells(
  tr: Element,
  replacedTokensMap: Map<Element, Set<string>>
): GridCellInfo[] {
  const tcs = directChildren(tr, 'tc');
  const cells: GridCellInfo[] = [];
  let colCursor = 0;

  for (const tc of tcs) {
    const tcPr = directChildren(tc, 'tcPr')[0] ?? null;
    const gridSpanEl = tcPr ? directChildren(tcPr, 'gridSpan')[0] : null;
    const spanRaw = gridSpanEl?.getAttribute('w:val') || gridSpanEl?.getAttribute('val') || '1';
    const gridSpan = Math.max(1, Number.parseInt(spanRaw, 10) || 1);

    const vMergeEl = tcPr ? directChildren(tcPr, 'vMerge')[0] : null;
    let vMerge: 'restart' | 'continue' | 'none' = 'none';
    if (vMergeEl) {
      const vmVal = vMergeEl.getAttribute('w:val') || vMergeEl.getAttribute('val') || '';
      vMerge = vmVal === 'restart' ? 'restart' : 'continue';
    }

    cells.push({
      tc,
      text: getCellPlainText(tc),
      hasTokens: cellReplacedAnyToken(replacedTokensMap, tc),
      colStart: colCursor,
      colEnd: colCursor + gridSpan,
      gridSpan,
      vMerge,
    });
    colCursor += gridSpan;
  }

  return cells;
}

function findGridCellAtCol(row: GridCellInfo[], targetCol: number): GridCellInfo | undefined {
  return row.find((c) => targetCol >= c.colStart && targetCol < c.colEnd);
}

function isBlankOrPlaceholderText(text: string): boolean {
  const t = text.trim();
  if (!t) return true;
  if (/^[._…\-–—/\s\[\]()]+$/.test(t)) return true;
  if (/^(?:<[^>]+>|\[[^\]]+\]|dd\.mm\.yyyy|to be filled|insert\b.*)$/i.test(t)) return true;
  return false;
}

function getStageDateForDocType(
  docType: MemoDocType,
  ph: Record<string, string>,
  stageMemo?: ResolvedStageMemo | null
): string | null {
  const si = stageMemo?.stageInputs;
  if (docType === 'mec') {
    return nonEmpty(si?.mecMemoDate) || nonEmpty(si?.icMemoDate) || nonEmpty(ph['#1#']);
  }
  if (docType === 'board') {
    return nonEmpty(si?.boardMemoDate) || nonEmpty(si?.mecMemoDate) || nonEmpty(ph['#1#']);
  }
  if (docType === 'sanction') {
    return nonEmpty(si?.sanctionDate) || nonEmpty(si?.boardMemoDate) || nonEmpty(ph['#1#']);
  }
  return nonEmpty(ph['#1#']);
}

function getStageMemoNoForDocType(
  docType: MemoDocType,
  stageMemo?: ResolvedStageMemo | null
): string | null {
  const si = stageMemo?.stageInputs;
  if (docType === 'mec') return nonEmpty(si?.mecMemoNumber);
  if (docType === 'board') return nonEmpty(si?.boardMemoNumber);
  if (docType === 'sanction') return nonEmpty(si?.sanctionRefNumber);
  return null;
}

function updateRemarksCellPreservingFormat(
  xmlDoc: Document,
  tcEl: Element,
  ph: Record<string, string>
) {
  const v = (code: string) => nonEmpty(ph[code]) || code;
  const paras = getCellParagraphs(tcEl);
  const deadlineVal = v('#28#');

  // Check if the template already structures Remarks as multiple separate styled paragraphs
  let updatedStructured = false;
  let expectingDeadlineNext = false;
  if (paras.length >= 3) {
    for (const pEl of paras) {
      const pTxt = getParagraphPlainText(pEl).trim();
      if (!pTxt) {
        if (expectingDeadlineNext) {
          setParagraphTextPreservingStyle(xmlDoc, pEl, deadlineVal);
          expectingDeadlineNext = false;
          updatedStructured = true;
        }
        continue;
      }
      if (/^1\s*\)\s*Settlement\s*Deadline\s*:?/i.test(pTxt)) {
        const afterColon = pTxt.split(':')[1]?.trim() ?? '';
        if (afterColon && !/^2\s*\)/.test(afterColon)) {
          setParagraphValueAfterColon(xmlDoc, pEl, deadlineVal);
          updatedStructured = true;
          expectingDeadlineNext = false;
        } else {
          expectingDeadlineNext = true;
        }
        continue;
      }
      if (expectingDeadlineNext) {
        if (/^2\s*\)/.test(pTxt)) {
          // Insert deadline paragraph right before "2) Break up"
          const cloned = pEl.cloneNode(true) as Element;
          setParagraphTextPreservingStyle(xmlDoc, cloned, deadlineVal);
          pEl.parentNode?.insertBefore(cloned, pEl);
          updatedStructured = true;
        } else {
          setParagraphTextPreservingStyle(xmlDoc, pEl, deadlineVal);
          updatedStructured = true;
        }
        expectingDeadlineNext = false;
        continue;
      }
      if (/^Principal\s*:/i.test(pTxt)) {
        setParagraphValueAfterColon(xmlDoc, pEl, v('#7#'));
        updatedStructured = true;
      } else if (/^Charges\s*:/i.test(pTxt)) {
        setParagraphValueAfterColon(xmlDoc, pEl, v('#8#'));
        updatedStructured = true;
      } else if (/^COF\s*:/i.test(pTxt)) {
        setParagraphValueAfterColon(xmlDoc, pEl, v('#31#'));
        updatedStructured = true;
      } else if (/^Total\s*:/i.test(pTxt)) {
        setParagraphValueAfterColon(xmlDoc, pEl, v('#32#'));
        updatedStructured = true;
      } else if (/^MV\s*:/i.test(pTxt)) {
        setParagraphValueAfterColon(xmlDoc, pEl, v('#33#'));
        updatedStructured = true;
      } else if (/^FSV\s*:/i.test(pTxt)) {
        setParagraphValueAfterColon(xmlDoc, pEl, v('#34#'));
        updatedStructured = true;
      }
    }
  }

  if (updatedStructured) {
    ensureCellWrapsAndFitsWithoutClipping(tcEl, `1)Settlement Deadline:\n${deadlineVal}\n2) Break up`);
    return;
  }

  const remarks = `1)Settlement Deadline:\n${deadlineVal}\n\n2) Break up\nPrincipal: ${v('#7#')}\nCharges: ${v('#8#')}\nCOF: ${v('#31#')}\nTotal: ${v('#32#')}\n\n3) Security Value:\nMV: ${v('#33#')}\nFSV: ${v('#34#')}`;
  setCellTextPreservingFormat(xmlDoc, tcEl, remarks);
}

function tryPopulateMetadataTable(
  xmlDoc: Document,
  matrix: GridCellInfo[][],
  docType: MemoDocType,
  ph: Record<string, string>,
  stageMemo?: ResolvedStageMemo | null
): boolean {
  const tableText = matrix.flat().map((c) => c.text).join(' | ');
  const looksLikeMetadata =
    (/\b(?:Dated|Date)\b/i.test(tableText) && /\b(?:Memo\s*no|Chronologic\s+Order|Sanction\s*Ref|To\b)/i.test(tableText)) ||
    /\b(?:IC|MEC|Board)\s+Memo\s+(?:Number|No|Date)\b/i.test(tableText) ||
    /\bSanction\s*(?:Ref\.?\s*No\.?|Number|Date)\b/i.test(tableText);
  if (!looksLikeMetadata) return false;

  const si = stageMemo?.stageInputs;
  const stageDate = getStageDateForDocType(docType, ph, stageMemo);
  const stageMemoNo = getStageMemoNoForDocType(docType, stageMemo);

  for (const row of matrix) {
    if (row.length < 2) continue;
    const valCell = row[row.length - 1];
    if (!valCell || valCell.hasTokens) continue;

    let labelCell = row[row.length - 2];
    if (labelCell.text.trim() === ':' && row.length >= 3) {
      labelCell = row[row.length - 3];
    }
    const label = labelCell.text.replace(/\s*:\s*$/, '').replace(/\s+/g, ' ').trim();
    if (!label) continue;

    if (/^(?:Chronologic\s+Order\s+of\s+(?:the\s+)?IC\s*Memo|IC\s*Memo)\s*(?:&|and)\s*Date$/i.test(label)) {
      if (nonEmpty(si?.icMemoNumber) || nonEmpty(si?.icMemoDate)) {
        setCellTextPreservingFormat(xmlDoc, valCell.tc, `${si?.icMemoNumber || '—'} dated ${si?.icMemoDate || '—'}`);
      }
    } else if (/^(?:Chronologic\s+Order\s+of\s+(?:the\s+)?MEC\s*Memo|MEC\s*Memo)\s*(?:&|and)\s*Date$/i.test(label)) {
      if (nonEmpty(si?.mecMemoNumber) || nonEmpty(si?.mecMemoDate)) {
        setCellTextPreservingFormat(xmlDoc, valCell.tc, `${si?.mecMemoNumber || '—'} dated ${si?.mecMemoDate || '—'}`);
      }
    } else if (/^(?:Chronologic\s+Order\s+of\s+(?:the\s+)?Board\s*Memo|Board\s*Memo)\s*(?:&|and)\s*Date$/i.test(label)) {
      if (nonEmpty(si?.boardMemoNumber) || nonEmpty(si?.boardMemoDate)) {
        setCellTextPreservingFormat(xmlDoc, valCell.tc, `${si?.boardMemoNumber || '—'} dated ${si?.boardMemoDate || '—'}`);
      }
    } else if (/^Prior\s+IC\s*(?:&|and)\s*MEC\s+Chronologic\s+Orders?$/i.test(label)) {
      if (nonEmpty(si?.icMemoNumber) || nonEmpty(si?.mecMemoNumber)) {
        setCellTextPreservingFormat(
          xmlDoc,
          valCell.tc,
          `IC: ${si?.icMemoNumber || '—'} (${si?.icMemoDate || '—'}) | MEC: ${si?.mecMemoNumber || '—'} (${si?.mecMemoDate || '—'})`
        );
      }
    } else if (/^(?:Chronologic\s+Order\s+of\s+(?:the\s+)?IC\s*Memo|IC\s*Memo\s*(?:Number|No\.?))$/i.test(label)) {
      if (nonEmpty(si?.icMemoNumber)) setCellTextPreservingFormat(xmlDoc, valCell.tc, si!.icMemoNumber);
    } else if (/^IC\s*Memo\s*Date$/i.test(label)) {
      if (nonEmpty(si?.icMemoDate)) setCellTextPreservingFormat(xmlDoc, valCell.tc, si!.icMemoDate);
    } else if (/^(?:Chronologic\s+Order\s+of\s+(?:the\s+)?MEC\s*Memo|MEC\s*Memo\s*(?:Number|No\.?))$/i.test(label)) {
      if (nonEmpty(si?.mecMemoNumber)) setCellTextPreservingFormat(xmlDoc, valCell.tc, si!.mecMemoNumber);
    } else if (/^MEC\s*Memo\s*Date$/i.test(label)) {
      if (nonEmpty(si?.mecMemoDate)) setCellTextPreservingFormat(xmlDoc, valCell.tc, si!.mecMemoDate);
    } else if (/^(?:Chronologic\s+Order\s+of\s+(?:the\s+)?Board\s*Memo|Board\s*Memo\s*(?:Number|No\.?))$/i.test(label)) {
      if (nonEmpty(si?.boardMemoNumber)) setCellTextPreservingFormat(xmlDoc, valCell.tc, si!.boardMemoNumber);
    } else if (/^Board\s*Memo\s*Date$/i.test(label)) {
      if (nonEmpty(si?.boardMemoDate)) setCellTextPreservingFormat(xmlDoc, valCell.tc, si!.boardMemoDate);
    } else if (/^Sanction\s*(?:Ref\.?\s*No\.?|Reference\s*No\.?|Number)$/i.test(label)) {
      if (nonEmpty(si?.sanctionRefNumber)) setCellTextPreservingFormat(xmlDoc, valCell.tc, si!.sanctionRefNumber);
    } else if (/^Sanction\s*Date$/i.test(label)) {
      if (nonEmpty(si?.sanctionDate)) setCellTextPreservingFormat(xmlDoc, valCell.tc, si!.sanctionDate);
    } else if (/^(?:Memo\s*(?:No\.?|Number)|Chronologic\s+Order|Ref\.?\s*No\.?)$/i.test(label)) {
      if (stageMemoNo) setCellTextPreservingFormat(xmlDoc, valCell.tc, stageMemoNo);
    } else if (/^(?:Dated|Date|Memo\s*Date)$/i.test(label)) {
      if (stageDate) setCellTextPreservingFormat(xmlDoc, valCell.tc, stageDate);
    } else if (/^To$/i.test(label) && docType === 'sanction' && isBlankOrPlaceholderText(valCell.text)) {
      if (nonEmpty(ph['#3#'])) {
        setCellTextPreservingFormat(xmlDoc, valCell.tc, `The Head of Branch, ${ph['#3#']} Branch`);
      }
    }
  }
  return true;
}

/**
 * Safely populates untokenized metadata tables or blank ref/date lines in headers/footers without
 * touching bank headings, logos, or Word PAGE/NUMPAGES fields.
 */
function applyHeaderFooterMetadata(
  xmlDoc: Document,
  memo: ResolvedICMemo,
  docType: MemoDocType,
  replacedTokensMap: Map<Element, Set<string>>,
  stageMemo?: ResolvedStageMemo | null
) {
  const ph = memo.placeholders;
  const tables = Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'tbl'));
  for (const tbl of tables) {
    const trs = directChildren(tbl, 'tr');
    const matrix = trs.map((tr) => getRowGridCells(tr, replacedTokensMap));
    tryPopulateMetadataTable(xmlDoc, matrix, docType, ph, stageMemo);
  }

  const stageDate = getStageDateForDocType(docType, ph, stageMemo);
  const stageMemoNo = getStageMemoNoForDocType(docType, stageMemo);
  const allParas = Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'p'));
  for (const pEl of allParas) {
    if (pEl.parentElement?.localName === 'tc') continue;
    if (firstDescendant(pEl, 'fldChar') || firstDescendant(pEl, 'fldSimple') || firstDescendant(pEl, 'instrText')) {
      continue;
    }
    const raw = getParagraphPlainText(pEl).trim();
    const colonIdx = raw.indexOf(':');
    if (colonIdx === -1) continue;
    const label = raw.slice(0, colonIdx).replace(/\s+/g, ' ').trim();
    const tail = raw.slice(colonIdx + 1).trim();
    if (!isBlankOrPlaceholderText(tail)) continue;
    if (/^(?:Sanction\s*(?:Ref\.?\s*No\.?|No\.?)|Memo\s*No\.?|Ref\.?\s*No\.?)$/i.test(label) && stageMemoNo) {
      setParagraphValueAfterColon(xmlDoc, pEl, stageMemoNo);
    } else if (/^(?:Dated|Date|Sanction\s*Date|Memo\s*Date)$/i.test(label) && stageDate) {
      setParagraphValueAfterColon(xmlDoc, pEl, stageDate);
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
  docType: MemoDocType,
  replacedTokensMap: Map<Element, Set<string>>,
  stageMemo?: ResolvedStageMemo | null,
  templateCtx?: TemplateStyleContext | null
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
    if (!normText || /\.{3,}\s*\d+$/.test(normText)) continue;

    // Inline Sanction/Stage top-level blank metadata paragraphs (only if outside multi-col tables and value is blank/placeholder)
    if (
      pEl.parentElement?.localName !== 'tc' &&
      !firstDescendant(pEl, 'fldChar') &&
      !firstDescendant(pEl, 'fldSimple') &&
      normText.includes(':')
    ) {
      const cIdx = text.indexOf(':');
      const inlineLabel = text.slice(0, cIdx).replace(/\s+/g, ' ').trim();
      const inlineTail = text.slice(cIdx + 1).trim();
      if (isBlankOrPlaceholderText(inlineTail)) {
        const stageDate = getStageDateForDocType(docType, ph, stageMemo);
        const stageMemoNo = getStageMemoNoForDocType(docType, stageMemo);
        if (/^(?:Sanction\s*(?:Ref\.?\s*No\.?|Reference\s*No\.?|No\.?)|Memo\s*No\.?|Ref\.?\s*No\.?)$/i.test(inlineLabel) && stageMemoNo) {
          setParagraphValueAfterColon(xmlDoc, pEl, stageMemoNo, null, templateCtx);
          continue;
        }
        if (/^(?:Dated|Date|Sanction\s*Date|Memo\s*Date)$/i.test(inlineLabel) && stageDate) {
          setParagraphValueAfterColon(xmlDoc, pEl, stageDate, null, templateCtx);
          continue;
        }
      }
    }

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
        populateNarrativeHeading(xmlDoc, pEl, ph['#54#'], false, templateCtx);
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
        setParagraphValueAfterColon(xmlDoc, pEl, bp.section5StatusNote, null, templateCtx);
      }
      continue;
    }

    // 6. Audit, Compliance header note
    if (/^(?:6\.\s*)?Audit,\s*Compliance[^:]*:/i.test(normText) && bp?.section6StatusNote !== undefined) {
      if (nonEmpty(bp.section6StatusNote)) {
        setParagraphValueAfterColon(xmlDoc, pEl, bp.section6StatusNote, null, templateCtx);
      }
      continue;
    }

    // 7a / 7b / 7c Notes & Collateral Schedules
    if (/^a\.\s*Primary\s*:/i.test(normText) && nonEmpty(bp?.section7PrimaryNote)) {
      setParagraphValueAfterColon(xmlDoc, pEl, bp!.section7PrimaryNote, null, templateCtx);
      continue;
    }
    if (/^b\.\s*Collateral\s*:/i.test(normText)) {
      if (nonEmpty(bp?.section7CollateralNote)) {
        setParagraphValueAfterColon(xmlDoc, pEl, bp!.section7CollateralNote, null, templateCtx);
      }
      if (bp?.section7CollateralSchedules && bp.section7CollateralSchedules.length > 0) {
        const subParas = collectContiguousSubParagraphs(pEl);
        let schedIdx = 0;
        for (const subP of subParas) {
          if (schedIdx >= bp.section7CollateralSchedules.length) break;
          const sTxt = getParagraphPlainText(subP).trim();
          if (/^c\.\s*Others/i.test(sTxt)) break;
          if (/^Schedule[-\s]*\d+/i.test(sTxt)) {
            setParagraphTextPreservingStyle(xmlDoc, subP, bp.section7CollateralSchedules[schedIdx], undefined, templateCtx);
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
        populateSubParagraphLines(xmlDoc, anchor, targetSub, customLines, '320', templateCtx);
      }
      continue;
    }

    // 8. Source of Repayment:
    if (
      /^(?:8\s*[.)]\s*)?Source\s+of\s+Repayment\s*:?/i.test(normText) &&
      nonEmpty(memo.sourceOfRepayment) &&
      !paraReplacedAnyOf(replacedTokensMap, pEl, [
        '#sourceOfRepayment#',
        '#sec8SourceOfRepayment#',
        '{{sourceOfRepayment}}',
        '[sourceOfRepayment]',
      ])
    ) {
      populateNarrativeHeading(xmlDoc, pEl, memo.sourceOfRepayment, false, templateCtx);
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
      populateNarrativeHeading(xmlDoc, pEl, memo.clientPrayer, false, templateCtx);
      continue;
    }

    // 12. Branch Visit Report:
    if (
      /^(?:\d+\s*[.)]\s*)?Branch\s+Visit\s+Report\s*:?/i.test(normText) &&
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
      populateNarrativeHeading(xmlDoc, pEl, memo.branchVisitReport, false, templateCtx);
      continue;
    }

    // 13. Recommendation of The Branch:
    if (/^(?:\d+\s*[.)]\s*)?(?:Recommendation\s+of\s+(?:the\s+)?Branch|Branch\s+Recommendation)\b/i.test(normText)) {
      const updateBranchRefAndDate = (targetPara: Element) => {
        if (
          nonEmpty(ph['#64#']) &&
          !paraReplacedAnyOf(replacedTokensMap, targetPara, ['#64#', '#refNo#', '#referenceNo#'])
        ) {
          const curTxt = getParagraphPlainText(targetPara);
          const refMatch = curTxt.match(/(Proposal\s*Ref\.?\s*:\s*)(.*?)(\s+dated\b)/i);
          if (refMatch && refMatch.index !== undefined) {
            const valStart = refMatch.index + refMatch[1].length;
            if (refMatch[2].length > 0) {
              replaceSpanInParagraph(xmlDoc, targetPara, valStart, valStart + refMatch[2].length, ph['#64#']);
            } else {
              // Insert right before " dated" by replacing the leading boundary char cleanly
              const spaceBefore = refMatch[1].endsWith(' ') ? ' ' : '';
              const repStart = Math.max(0, valStart - spaceBefore.length);
              replaceSpanInParagraph(xmlDoc, targetPara, repStart, valStart, `${spaceBefore}${ph['#64#']}`);
            }
          }
        }
        if (
          nonEmpty(ph['#65#']) &&
          !paraReplacedAnyOf(replacedTokensMap, targetPara, ['#65#', '#branchProposalDate#'])
        ) {
          const curTxt = getParagraphPlainText(targetPara);
          const dateMatch = curTxt.match(/(\bdated\s*)(.*)$/i);
          if (dateMatch && dateMatch.index !== undefined) {
            const dateStart = dateMatch.index + dateMatch[1].length;
            if (dateMatch[2].trim().length > 0) {
              replaceSpanInParagraph(xmlDoc, targetPara, dateStart, curTxt.length, ph['#65#']);
            } else {
              replaceSpanInParagraph(xmlDoc, targetPara, dateMatch.index, curTxt.length, `dated ${ph['#65#']}`);
            }
          }
        }
      };

      updateBranchRefAndDate(pEl);

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
          const firstSub = subParas[0];
          const firstSubNorm = firstSub ? getParagraphPlainText(firstSub).replace(/\s+/g, ' ').trim() : '';
          if (firstSub && /^The\s+Branch\s+Recommended/i.test(firstSubNorm)) {
            updateBranchRefAndDate(firstSub);
            const subRefRPr =
              (templateCtx ? getParagraphOwnBodyRPr(xmlDoc, templateCtx.styles, firstSub) : null) ??
              resolveSectionBodyRPr(xmlDoc, templateCtx, pEl, subParas.slice(1));
            populateSubParagraphLines(
              xmlDoc,
              firstSub,
              subParas.slice(1),
              recLines,
              '180',
              templateCtx,
              subRefRPr
            );
          } else if (/The\s+Branch\s+Recommended/i.test(normText) || subParas.length > 0) {
            populateSubParagraphLines(xmlDoc, pEl, subParas, recLines, '180', templateCtx);
          } else {
            populateNarrativeHeading(xmlDoc, pEl, recLines.join('\n'), false, templateCtx);
          }
        }
      }
      continue;
    }

    // 14. Recommendation of Internal Control & Compliance Division / ICCD:
    if (
      /^(?:\d+\s*[.)]\s*)?Recommendation\s+of\s+(?:Internal\s+Control|IC\s*&?\s*CD|ICCD)\b/i.test(normText) &&
      nonEmpty(memo.sec14IccdRecommendation) &&
      !paraReplacedAnyOf(replacedTokensMap, pEl, [
        '#iccdRecommendation#',
        '#sec14IccdRecommendation#',
        '{{iccdRecommendation}}',
        '{{sec14IccdRecommendation}}',
      ])
    ) {
      populateNarrativeHeading(xmlDoc, pEl, memo.sec14IccdRecommendation, false, templateCtx);
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
              setParagraphTextPreservingStyle(xmlDoc, aBulletPara, `${prefix}${cleanLines.join('\n')}`, undefined, templateCtx);
            }
            // Remove extra stale continuation lines between "a." and "b."
            for (let k = aIdx + 1; k < subParas.length; k++) {
              subParas[k].parentNode?.removeChild(subParas[k]);
            }
          } else {
            const colonIdx = text.indexOf(':');
            if (subParas.length === 0 && colonIdx !== -1 && text.slice(colonIdx + 1).trim().length > 0) {
              setParagraphValueAfterColon(xmlDoc, pEl, cleanLines.join('\n'), null, templateCtx);
            } else {
              const formattedLines = cleanLines.map((l, idx) => (idx === 0 ? `a. ${l}` : l));
              populateSubParagraphLines(xmlDoc, pEl, subParas, formattedLines, '120', templateCtx);
            }
          }
        }
      }
      continue;
    }

    // 15. Observation and Recommendation of SAMD:
    if (
      /^(?:\d+\s*[.)]\s*)?(?:Observations?\s*(?:and|&|\/)\s*Recommendations?\s*of\s*SAMD|SAMD\s*Observations?(?:\s*(?:and|&|\/)\s*Recommendations?)?)\s*:?/i.test(
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
          populateNarrativeHeading(xmlDoc, pEl, obsLines.join('\n'), true, templateCtx);
        } else {
          // Clear any stale inline text after colon if SAMD points live in sub-paragraphs
          const colonIdx = text.indexOf(':');
          if (colonIdx !== -1 && text.slice(colonIdx + 1).trim().length > 0) {
            setParagraphValueAfterColon(xmlDoc, pEl, '', null, templateCtx);
          }
          populateSubParagraphLines(xmlDoc, pEl, subParas, obsLines, '180', templateCtx);
        }
      }
      continue;
    }

    // Previous Board Decision(s):
    if (
      /^Previous\s+Board\s+Decision(?:\(s\))?[^:]*:/i.test(normText) &&
      nonEmpty(bp?.previousBoardDecisions) &&
      !paraReplacedAnyOf(replacedTokensMap, pEl, ['#previousBoardDecisions#', '{{previousBoardDecisions}}'])
    ) {
      setParagraphValueAfterColon(xmlDoc, pEl, bp!.previousBoardDecisions, null, templateCtx);
      continue;
    }

    // MEC / Board / Sanction Stage Sections
    if (stageMemo) {
      const ss = stageMemo.stageSections;

      // 16. Recommendation of the Investment Committee (IC)
      if (
        /^(?:\d+\s*[.)]\s*)?(?:Recommendation\s+of\s+(?:the\s+)?Investment\s+Committee(?:\s*\(IC\))?|Investment\s+Committee\s*(?:\(IC\))?\s+Recommendation)(?:\s*\(Carried\s+Forward\))?\s*:?/i.test(
          normText
        ) &&
        nonEmpty(ss.mecIcRecommendation)
      ) {
        const subParas = collectContiguousSubParagraphs(pEl);
        const alreadyTokenized =
          paraReplacedAnyOf(replacedTokensMap, pEl, ['#mecIcRecommendation#', '{{mecIcRecommendation}}', '[mecIcRecommendation]']) ||
          subParas.some((sp) =>
            paraReplacedAnyOf(replacedTokensMap, sp, ['#mecIcRecommendation#', '{{mecIcRecommendation}}', '[mecIcRecommendation]'])
          );
        if (!alreadyTokenized) {
          populateNarrativeHeading(xmlDoc, pEl, ss.mecIcRecommendation, true, templateCtx);
        }
        continue;
      }

      // 17. Management Executive Committee (MEC) Recommendation
      if (
        /^(?:\d+\s*[.)]\s*)?(?:Recommendation\s+of\s+(?:the\s+)?Management\s+Executive\s+Committee(?:\s*\(MEC\))?|Management\s+Executive\s+Committee\s*(?:\(MEC\))?\s+Recommendation)(?:\s*\(Carried\s+Forward\))?\s*:?/i.test(
          normText
        ) &&
        nonEmpty(ss.boardMecRecommendation)
      ) {
        const subParas = collectContiguousSubParagraphs(pEl);
        const alreadyTokenized =
          paraReplacedAnyOf(replacedTokensMap, pEl, ['#boardMecRecommendation#', '{{boardMecRecommendation}}', '[boardMecRecommendation]']) ||
          subParas.some((sp) =>
            paraReplacedAnyOf(replacedTokensMap, sp, ['#boardMecRecommendation#', '{{boardMecRecommendation}}', '[boardMecRecommendation]'])
          );
        if (!alreadyTokenized) {
          populateNarrativeHeading(xmlDoc, pEl, ss.boardMecRecommendation, true, templateCtx);
        }
        continue;
      }

      // Sanction: AI Cross-Verified Facility & Liability Confirmation
      if (
        /^(?:\d+\s*[.)]\s*)?AI\s+Cross-Verified\s+Facility(?:\s*(?:&|and)\s*Liability\s+Confirmation)?\s*:?/i.test(
          normText
        ) &&
        nonEmpty(ss.sanctionCrossVerifiedSummary)
      ) {
        const subParas = collectContiguousSubParagraphs(pEl);
        const alreadyTokenized =
          paraReplacedAnyOf(replacedTokensMap, pEl, ['#sanctionCrossVerifiedSummary#', '{{sanctionCrossVerifiedSummary}}', '[sanctionCrossVerifiedSummary]']) ||
          subParas.some((sp) =>
            paraReplacedAnyOf(replacedTokensMap, sp, ['#sanctionCrossVerifiedSummary#', '{{sanctionCrossVerifiedSummary}}', '[sanctionCrossVerifiedSummary]'])
          );
        if (!alreadyTokenized) {
          populateNarrativeHeading(xmlDoc, pEl, ss.sanctionCrossVerifiedSummary, true, templateCtx);
        }
        continue;
      }
    }
  }

  // 2. Process all tables in document (preserving 100% of original table layout, borders, cell widths, margins, and row properties)
  const tables = Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'tbl'));
  for (const tbl of tables) {
    const trs = directChildren(tbl, 'tr');
    if (trs.length === 0) continue;

    const matrix = trs.map((tr) => getRowGridCells(tr, replacedTokensMap));
    const tableText = matrix.flat().map((c) => c.text).join(' | ');

    // A. Metadata Table (Dated / Memo no / To / From + Stage Memo Numbers & Dates)
    if (tryPopulateMetadataTable(xmlDoc, matrix, docType, ph, stageMemo)) {
      continue;
    }

    // B. A/C & Branch + Subject Bar (preserve mixed bold/non-bold run styles inside cells)
    if (/\bA\/C\s*:/i.test(tableText) && /\bBranch\s*:/i.test(tableText)) {
      for (const row of matrix) {
        for (const c of row) {
          if (c.hasTokens) continue;
          const paras = getCellParagraphs(c.tc);
          if (/\bA\/C\s*:/i.test(c.text) && nonEmpty(ph['#2#'])) {
            const cis = nonEmpty(ph['#69#']) || '#69#';
            const targetP = paras.find((p) => /\bA\/C\s*:/i.test(getParagraphPlainText(p))) ?? paras[0];
            if (targetP) {
              const pTxt = getParagraphPlainText(targetP);
              const m = pTxt.match(/(A\/C\s*:\s*)(.+?)(\s*\(\s*CIS\s*:\s*)([^)]+)(\))/i);
              if (m) {
                const acStart = pTxt.indexOf(m[2], pTxt.toLowerCase().indexOf('a/c'));
                if (acStart !== -1) {
                  replaceSpanInParagraph(xmlDoc, targetP, acStart, acStart + m[2].length, ph['#2#']);
                }
                const refreshed = getParagraphPlainText(targetP);
                const mCis = refreshed.match(/(\(\s*CIS\s*:\s*)([^)]+)(\))/i);
                if (mCis) {
                  const cisStart = refreshed.indexOf(mCis[2], refreshed.toLowerCase().indexOf('cis'));
                  if (cisStart !== -1) {
                    replaceSpanInParagraph(xmlDoc, targetP, cisStart, cisStart + mCis[2].length, cis);
                  }
                }
              } else if (!setParagraphValueAfterColon(xmlDoc, targetP, `${ph['#2#']} (CIS: ${cis})`)) {
                setCellTextPreservingFormat(xmlDoc, c.tc, `A/C: ${ph['#2#']} (CIS: ${cis})`);
              }
            }
          } else if (/\bBranch\s*:/i.test(c.text) && nonEmpty(ph['#3#'])) {
            const targetP = paras.find((p) => /\bBranch\s*:/i.test(getParagraphPlainText(p))) ?? paras[0];
            if (!targetP || !setParagraphValueAfterColon(xmlDoc, targetP, ph['#3#'])) {
              setCellTextPreservingFormat(xmlDoc, c.tc, `Branch: ${ph['#3#']}`);
            }
          } else if (/^Subject\s*:/i.test(c.text) && nonEmpty(ph['#4#'])) {
            const targetP = paras.find((p) => /^Subject\s*:/i.test(getParagraphPlainText(p).trim())) ?? paras[0];
            if (targetP) {
              const pTxt = getParagraphPlainText(targetP);
              const m = pTxt.match(/waiver\s+of\s+Tk\.?\s*([^\s,]+)/i);
              if (m) {
                const start = pTxt.indexOf(m[1], pTxt.toLowerCase().search(/waiver\s+of\s+tk/));
                if (start !== -1) {
                  replaceSpanInParagraph(xmlDoc, targetP, start, start + m[1].length, ph['#4#']);
                }
              }
            }
          }
        }
      }
      continue;
    }

    // C. Main Settlement Table (Particulars + Grand-Total + Principal) — vMerge & gridSpan aware
    if (/\bPrincipal\b/i.test(tableText) && /\bGrand[-\s]*Total\b/i.test(tableText)) {
      let osColIdx = 4;
      for (let r = 0; r < matrix.length; r++) {
        const row = matrix[r];
        const joined = row.map((c) => c.text).join(' | ');

        // Header row "O/S as on ..."
        for (const c of row) {
          if (!c.hasTokens && /O\/S\s+as\s+on\b/i.test(c.text) && nonEmpty(ph['#68#'])) {
            const targetP = getCellParagraphs(c.tc).find((p) => /O\/S\s+as\s+on\b/i.test(getParagraphPlainText(p)));
            if (targetP) {
              const pTxt = getParagraphPlainText(targetP);
              const m = pTxt.match(/(O\/S\s+as\s+on\s+)(.+)/i);
              if (m) {
                const st = pTxt.indexOf(m[2], pTxt.toLowerCase().search(/o\/s\s+as\s+on/));
                if (st !== -1) {
                  replaceSpanInParagraph(xmlDoc, targetP, st, st + m[2].length, ph['#68#']);
                }
              }
            } else {
              const updated = c.text.replace(/(O\/S\s+as\s+on\s+).*/i, `$1${ph['#68#']}`);
              setCellTextPreservingFormat(xmlDoc, c.tc, updated);
            }
          }
        }

        // First data row containing Principal (also contains Mode, Limit Original, RSD, Remarks via rowSpan/vMerge)
        const princCell = row.find((c) => c.vMerge !== 'continue' && /^Principal$/i.test(c.text.trim()));
        if (princCell) {
          osColIdx = princCell.colEnd;
          if (princCell.colStart >= 3) {
            const modeCell = findGridCellAtCol(row, 0);
            const origCell = findGridCellAtCol(row, 1);
            const rsdCell = findGridCellAtCol(row, 2);
            if (modeCell && !modeCell.hasTokens && nonEmpty(ph['#5#'])) {
              setCellTextPreservingFormat(xmlDoc, modeCell.tc, ph['#5#']);
            }
            if (origCell && !origCell.hasTokens && nonEmpty(ph['#6#'])) {
              setCellTextPreservingFormat(xmlDoc, origCell.tc, ph['#6#']);
            }
            if (rsdCell && !rsdCell.hasTokens && nonEmpty(bp?.limitRSD)) {
              setCellTextPreservingFormat(xmlDoc, rsdCell.tc, bp!.limitRSD);
            }
          }
          const osCell = findGridCellAtCol(row, osColIdx);
          if (osCell && !osCell.hasTokens && nonEmpty(ph['#7#'])) {
            setCellTextPreservingFormat(xmlDoc, osCell.tc, ph['#7#']);
          }
          const remarksCell = row[row.length - 1];
          if (
            remarksCell &&
            !remarksCell.hasTokens &&
            /Settlement\s*Deadline|Break\s*up|Security\s*Value/i.test(remarksCell.text)
          ) {
            updateRemarksCellPreservingFormat(xmlDoc, remarksCell.tc, ph);
          }
          continue;
        }

        const labelCell = row.find((c) => c.vMerge !== 'continue' && c.text.trim() !== '');
        if (labelCell) {
          const lblTxt = labelCell.text.trim();
          const valCell = findGridCellAtCol(row, labelCell.colEnd);
          if (/^Charges\b/i.test(lblTxt) && valCell && !valCell.hasTokens && nonEmpty(ph['#8#'])) {
            setCellTextPreservingFormat(xmlDoc, valCell.tc, ph['#8#']);
            continue;
          }
          if (/^Charged\s*Profit$/i.test(lblTxt) && valCell && !valCell.hasTokens && nonEmpty(ph['#10#'])) {
            setCellTextPreservingFormat(xmlDoc, valCell.tc, ph['#10#']);
            continue;
          }
          if (/^Compensation$/i.test(lblTxt) && valCell && !valCell.hasTokens && nonEmpty(ph['#11#'])) {
            setCellTextPreservingFormat(xmlDoc, valCell.tc, ph['#11#']);
            continue;
          }
          if (/^Uncharged\s*Compensation$/i.test(lblTxt) && valCell && !valCell.hasTokens && nonEmpty(ph['#12#'])) {
            setCellTextPreservingFormat(xmlDoc, valCell.tc, ph['#12#']);
            continue;
          }
        }

        if (/Grand[-\s]*Total/i.test(joined)) {
          const gtCell = row.find((c) => /Grand[-\s]*Total/i.test(c.text));
          const gtStartCol = gtCell ? Math.max(gtCell.colEnd, osColIdx) : osColIdx;
          const osTotal = findGridCellAtCol(row, gtStartCol);
          const waiverTotal = findGridCellAtCol(row, gtStartCol + 1);
          const settleTotal = findGridCellAtCol(row, gtStartCol + 2);
          if (osTotal && !osTotal.hasTokens && nonEmpty(ph['#25#'])) {
            setCellTextPreservingFormat(xmlDoc, osTotal.tc, ph['#25#']);
          }
          if (waiverTotal && !waiverTotal.hasTokens && nonEmpty(ph['#4#'])) {
            setCellTextPreservingFormat(xmlDoc, waiverTotal.tc, ph['#4#']);
          }
          if (settleTotal && !settleTotal.hasTokens && nonEmpty(ph['#27#'])) {
            setCellTextPreservingFormat(xmlDoc, settleTotal.tc, ph['#27#']);
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

    // E. Key Points Table (handle inline ": val", blank cell, or dedicated ":" separator column)
    if (/Cost of Fund/i.test(tableText) && /Maximum Waiver/i.test(tableText)) {
      for (const row of matrix) {
        for (let ci = 0; ci < row.length - 1; ci++) {
          if (row[ci].vMerge === 'continue') continue;
          const label = row[ci].text.replace(/\s*:\s*$/, '').trim();
          let nextIdx = ci + 1;
          let hasSeparateColonCell = false;
          if (row[nextIdx]?.text.trim() === ':' && nextIdx + 1 < row.length) {
            hasSeparateColonCell = true;
            nextIdx++;
          }
          const next = row[nextIdx];
          if (!next || next.hasTokens || next.vMerge === 'continue') continue;

          const writeKeyPoint = (val: string | null) => {
            if (!val) return;
            const formatted =
              hasSeparateColonCell || !next.text.trim().startsWith(':') && next.text.trim() !== '' && !isBlankOrPlaceholderText(next.text)
                ? val
                : `: ${val}`;
            setCellTextPreservingFormat(xmlDoc, next.tc, formatted);
          };

          if (/Total\s*\[Profit/i.test(label)) writeKeyPoint(nonEmpty(ph['#42#']));
          else if (/^Cost of Fund$/i.test(label)) writeKeyPoint(nonEmpty(ph['#31#']));
          else if (/Maximum Waiver/i.test(label)) writeKeyPoint(nonEmpty(ph['#44#']));
          else if (/^(Client|Branch)$/i.test(label)) writeKeyPoint(nonEmpty(ph['#45#']));
          else if (/^(SAMD,\s*HO|HOIC|MEC|IC\s*&\s*CD)$/i.test(label)) writeKeyPoint(nonEmpty(ph['#47#']));
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

    // G. Section 2 Liability Position Table (vMerge & gridSpan aware)
    if (/Name of the Concern/i.test(tableText) && /Mode of Investment/i.test(tableText)) {
      for (const row of matrix) {
        const modeCell = findGridCellAtCol(row, 1);
        const isFundedRow =
          Boolean(modeCell) &&
          /\bFunded\b/i.test(modeCell!.text) &&
          !/Non[-\s]*Funded/i.test(modeCell!.text);
        const isNonFundedRow =
          Boolean(modeCell) && /^Non[-\s]*Funded$/i.test(modeCell!.text.trim());
        const totalCell = row.find((c) => c.vMerge !== 'continue' && /^Total$/i.test(c.text.trim()));
        const isTotalRow = Boolean(totalCell) && row.length >= 6;

        if (isFundedRow) {
          const c0 = findGridCellAtCol(row, 0);
          const c1 = findGridCellAtCol(row, 1);
          const c2 = findGridCellAtCol(row, 2);
          const c3 = findGridCellAtCol(row, 3);
          const c4 = findGridCellAtCol(row, 4);
          const c5 = findGridCellAtCol(row, 5);
          const c6 = findGridCellAtCol(row, 6);
          const c7 = findGridCellAtCol(row, 7);
          if (c0 && !c0.hasTokens && nonEmpty(ph['#2#'])) setCellTextPreservingFormat(xmlDoc, c0.tc, ph['#2#']);
          if (c1 && !c1.hasTokens && nonEmpty(ph['#55#'])) setCellTextPreservingFormat(xmlDoc, c1.tc, `${ph['#55#']}\nFunded`);
          if (c2 && !c2.hasTokens && ph['#56#'] !== undefined && (nonEmpty(ph['#56#']) || bp !== null)) {
            setCellTextPreservingFormat(xmlDoc, c2.tc, ph['#56#'] || '');
          }
          if (c3 && !c3.hasTokens && nonEmpty(ph['#25#'])) setCellTextPreservingFormat(xmlDoc, c3.tc, ph['#25#']);
          if (c4 && !c4.hasTokens && ph['#58#'] !== undefined && (nonEmpty(ph['#58#']) || bp !== null)) {
            setCellTextPreservingFormat(xmlDoc, c4.tc, ph['#58#'] || '');
          }
          if (c5 && !c5.hasTokens && nonEmpty(ph['#59#'])) setCellTextPreservingFormat(xmlDoc, c5.tc, ph['#59#']);
          if (c6 && !c6.hasTokens && nonEmpty(ph['#60#'])) setCellTextPreservingFormat(xmlDoc, c6.tc, ph['#60#']);
          if (c7 && !c7.hasTokens && bp !== null) {
            setCellTextPreservingFormat(xmlDoc, c7.tc, bp.sec2Remarks || '');
          }
        } else if (isNonFundedRow) {
          const remarksCell = row[row.length - 1];
          if (remarksCell && !remarksCell.hasTokens && memo.sec2NonFundedNote !== undefined && (nonEmpty(memo.sec2NonFundedNote) || bp !== null)) {
            setCellTextPreservingFormat(xmlDoc, remarksCell.tc, memo.sec2NonFundedNote || '');
          }
        } else if (isTotalRow) {
          const startCol = Math.max(totalCell!.colEnd, 2);
          const tLim = findGridCellAtCol(row, startCol);
          const tOs = findGridCellAtCol(row, startCol + 1);
          const tVal = findGridCellAtCol(row, startCol + 2);
          const tOv = findGridCellAtCol(row, startCol + 3);
          const tRec = findGridCellAtCol(row, startCol + 4);
          if (tLim && !tLim.hasTokens && nonEmpty(ph['#56#'])) setCellTextPreservingFormat(xmlDoc, tLim.tc, ph['#56#']);
          if (tOs && !tOs.hasTokens && nonEmpty(ph['#25#'])) setCellTextPreservingFormat(xmlDoc, tOs.tc, ph['#25#']);
          if (tVal && !tVal.hasTokens && nonEmpty(ph['#58#'])) setCellTextPreservingFormat(xmlDoc, tVal.tc, ph['#58#']);
          if (tOv && !tOv.hasTokens && nonEmpty(ph['#59#'])) setCellTextPreservingFormat(xmlDoc, tOv.tc, ph['#59#']);
          if (tRec && !tRec.hasTokens && nonEmpty(ph['#60#'])) setCellTextPreservingFormat(xmlDoc, tRec.tc, ph['#60#']);
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
        let labelCell = row[row.length - 2];
        if (labelCell.text.trim() === ':' && row.length >= 3) {
          labelCell = row[row.length - 3];
        }
        const label = labelCell.text.replace(/\s*:\s*$/, '').replace(/\s+/g, ' ').trim();
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
    const stageDate = getStageDateForDocType(docType, ph, stageMemo);
    const stageMemoNo = getStageMemoNoForDocType(docType, stageMemo);
    for (const row of matrix) {
      if (row.length < 2 || row.length > 4) continue;
      const valCell = row[row.length - 1];
      if (!valCell || valCell.hasTokens || valCell.vMerge === 'continue') continue;

      // Determine label cell (skip ":" separator cell)
      let labelCell = row[row.length - 2];
      if (labelCell.text.trim() === ':' && row.length >= 3) {
        labelCell = row[row.length - 3];
      }
      const rawLabel = labelCell.text.trim();
      if (!rawLabel) continue;
      // Normalize optional leading serial ("a. ", "1. ") and trailing ":" + collapse whitespace
      const label = rawLabel
        .replace(/^(?:[a-z]\.|\d+[.)])\s+/i, '')
        .replace(/\s*:\s*$/, '')
        .replace(/\s+/g, ' ')
        .trim();
      if (!label) continue;

      const sp = bp?.section10Particulars;
      const trySet = (val: string | null) => {
        if (val !== null) {
          const lblParas = getCellParagraphs(labelCell.tc);
          const lblFallbackRPr = lblParas[0]
            ? buildCleanValueRPr(xmlDoc, lblParas[0], templateCtx)
            : templateCtx?.narrativeBodyRPr ?? null;
          setCellTextPreservingFormat(xmlDoc, valCell.tc, val, lblFallbackRPr, templateCtx);
        }
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
      else if (/^Proposal\s*Date$/i.test(label)) trySet(nonEmpty(ph['#1#']));
      else if (/^(Memo\s*Date|Dated)$/i.test(label)) trySet(stageDate);
      else if (/^(Memo\s*(?:No\.?|Number)|Chronologic\s+Order)$/i.test(label) && stageMemoNo) trySet(stageMemoNo);
      else if (/^(Proposed\s*Waiver|Waiver\s*Amount)$/i.test(label)) trySet(nonEmpty(ph['#4#']));
      else if (/^(Settlement\s*Amount|Exit\s*Amount)$/i.test(label)) trySet(nonEmpty(ph['#27#']));
      else if (/^(Settlement\s*Period|Settlement\s*Deadline)$/i.test(label)) trySet(nonEmpty(ph['#28#']));
      else if (/^(?:8\s*[.)]\s*)?Source\s+of\s+Repayment$/i.test(label)) trySet(nonEmpty(memo.sourceOfRepayment));
      else if (/^(?:12\s*[.)]\s*)?Contents\s+of\s+(?:the\s+)?Client[’']?s\s+Prayer$/i.test(label)) trySet(nonEmpty(memo.clientPrayer));
      else if (/^(?:\d+\s*[.)]\s*)?Branch\s+Visit\s+Report$/i.test(label)) trySet(nonEmpty(memo.branchVisitReport));
      else if (/^(?:\d+\s*[.)]\s*)?(?:Recommendation\s+of\s+(?:the\s+)?Branch|Branch\s+Recommendation)$/i.test(label)) trySet(nonEmpty(memo.sec13BranchRecommendation));
      else if (/^(?:\d+\s*[.)]\s*)?Recommendation\s+of\s+(?:Internal\s+Control|IC\s*&?\s*CD|ICCD)/i.test(label)) {
        trySet(nonEmpty(memo.sec14IccdRecommendation));
      } else if (/^(?:\d+\s*[.)]\s*)?(?:Observations?\s*(?:and|&|\/)\s*Recommendations?\s*of\s*SAMD|SAMD\s*Observations?(?:\s*(?:and|&|\/)\s*Recommendations?)?)$/i.test(label)) {
        trySet(nonEmpty(memo.samdObservation));
      } else if (/^Justification\s+of\s+(?:the\s+)?Proposed\s+Facility$/i.test(label)) trySet(nonEmpty(memo.justificationLines));
      else if (stageMemo) {
        const si = stageMemo.stageInputs;
        const ss = stageMemo.stageSections;
        if (/^(?:Chronologic\s+Order\s+of\s+(?:the\s+)?IC\s*Memo|IC\s*Memo)\s*(?:&|and)\s*Date$/i.test(label) && (nonEmpty(si.icMemoNumber) || nonEmpty(si.icMemoDate))) {
          trySet(`${si.icMemoNumber || '—'} dated ${si.icMemoDate || '—'}`);
        } else if (/^(?:Chronologic\s+Order\s+of\s+(?:the\s+)?MEC\s*Memo|MEC\s*Memo)\s*(?:&|and)\s*Date$/i.test(label) && (nonEmpty(si.mecMemoNumber) || nonEmpty(si.mecMemoDate))) {
          trySet(`${si.mecMemoNumber || '—'} dated ${si.mecMemoDate || '—'}`);
        } else if (/^(?:Chronologic\s+Order\s+of\s+(?:the\s+)?Board\s*Memo|Board\s*Memo)\s*(?:&|and)\s*Date$/i.test(label) && (nonEmpty(si.boardMemoNumber) || nonEmpty(si.boardMemoDate))) {
          trySet(`${si.boardMemoNumber || '—'} dated ${si.boardMemoDate || '—'}`);
        } else if (/^(?:Chronologic\s+Order\s+of\s+(?:the\s+)?IC\s*Memo|IC\s*Memo\s*(?:Number|No\.?))$/i.test(label)) trySet(nonEmpty(si.icMemoNumber));
        else if (/^IC\s*Memo\s*Date$/i.test(label)) trySet(nonEmpty(si.icMemoDate));
        else if (/^(?:Chronologic\s+Order\s+of\s+(?:the\s+)?MEC\s*Memo|MEC\s*Memo\s*(?:Number|No\.?))$/i.test(label)) trySet(nonEmpty(si.mecMemoNumber));
        else if (/^MEC\s*Memo\s*Date$/i.test(label)) trySet(nonEmpty(si.mecMemoDate));
        else if (/^(?:Chronologic\s+Order\s+of\s+(?:the\s+)?Board\s*Memo|Board\s*Memo\s*(?:Number|No\.?))$/i.test(label)) trySet(nonEmpty(si.boardMemoNumber));
        else if (/^Board\s*Memo\s*Date$/i.test(label)) trySet(nonEmpty(si.boardMemoDate));
        else if (/^Sanction\s*(?:Ref\.?\s*No\.?|Reference\s*No\.?|Number)$/i.test(label)) trySet(nonEmpty(si.sanctionRefNumber));
        else if (/^Sanction\s*Date$/i.test(label)) trySet(nonEmpty(si.sanctionDate));
        else if (/^(?:Recommendation\s+of\s+(?:the\s+)?Investment\s+Committee(?:\s*\(IC\))?|Investment\s+Committee\s*(?:\(IC\))?\s+Recommendation)(?:\s*\(Carried\s+Forward\))?$/i.test(label)) trySet(nonEmpty(ss.mecIcRecommendation));
        else if (/^(?:Recommendation\s+of\s+(?:the\s+)?Management\s+Executive\s+Committee(?:\s*\(MEC\))?|Management\s+Executive\s+Committee\s*(?:\(MEC\))?\s+Recommendation)(?:\s*\(Carried\s+Forward\))?$/i.test(label)) trySet(nonEmpty(ss.boardMecRecommendation));
        else if (/^AI\s+Cross-Verified\s+Facility/i.test(label)) trySet(nonEmpty(ss.sanctionCrossVerifiedSummary));
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
    populateNarrativeHeading(xmlDoc, targetP, trimmed, preferSub, templateCtx);
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
    /^(?:\d+\s*[.)]\s*)?Branch\s+Visit\s+Report\s*:?/i,
    false
  );
  ensureHeadingValue(
    memo.sec13BranchRecommendation,
    /^(?:\d+\s*[.)]\s*)?(?:Recommendation\s+of\s+(?:the\s+)?Branch|Branch\s+Recommendation)\b/i,
    true
  );
  ensureHeadingValue(
    memo.sec14IccdRecommendation,
    /^(?:\d+\s*[.)]\s*)?Recommendation\s+of\s+(?:Internal\s+Control|IC\s*&?\s*CD|ICCD)\b/i,
    false
  );
  ensureHeadingValue(
    memo.samdObservation,
    /^(?:\d+\s*[.)]\s*)?(?:Observations?\s*(?:and|&|\/)\s*Recommendations?\s*of\s*SAMD|SAMD\s*Observations?(?:\s*(?:and|&|\/)\s*Recommendations?)?)\b/i,
    true
  );

  if (stageMemo) {
    const ss = stageMemo.stageSections;
    ensureHeadingValue(
      ss.mecIcRecommendation,
      /^(?:\d+\s*[.)]\s*)?(?:Recommendation\s+of\s+(?:the\s+)?Investment\s+Committee(?:\s*\(IC\))?|Investment\s+Committee\s*(?:\(IC\))?\s+Recommendation)(?:\s*\(Carried\s+Forward\))?\s*:?/i,
      true
    );
    ensureHeadingValue(
      ss.boardMecRecommendation,
      /^(?:\d+\s*[.)]\s*)?(?:Recommendation\s+of\s+(?:the\s+)?Management\s+Executive\s+Committee(?:\s*\(MEC\))?|Management\s+Executive\s+Committee\s*(?:\(MEC\))?\s+Recommendation)(?:\s*\(Carried\s+Forward\))?\s*:?/i,
      true
    );
    ensureHeadingValue(
      ss.sanctionCrossVerifiedSummary,
      /^(?:\d+\s*[.)]\s*)?AI\s+Cross-Verified\s+Facility(?:\s*(?:&|and)\s*Liability\s+Confirmation)?\s*:?/i,
      true
    );
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
  docType: MemoDocType,
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

  const stylesFile = zip.file('word/styles.xml');
  const stylesDoc = stylesFile
    ? parser.parseFromString(await stylesFile.async('string'), 'application/xml')
    : null;
  const stylesInfo = parseStylesXml(stylesDoc);

  const xmlPaths = Object.keys(zip.files).filter((p) =>
    /^word\/(document|header\d*|footer\d*)\.xml$/i.test(p)
  );

  for (const xp of xmlPaths) {
    const origXmlStr = await zip.files[xp].async('string');
    const xmlDeclMatch = origXmlStr.match(/^<\?xml[^?]*\?>\s*/i);
    const xmlDoc = parser.parseFromString(origXmlStr, 'application/xml');

    const templateCtx =
      xp === 'word/document.xml'
        ? buildTemplateStyleContext(xmlDoc, stylesInfo)
        : null;

    const replacedTokensMap = new Map<Element, Set<string>>();
    const allParagraphs = Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'p'));

    for (const pEl of allParagraphs) {
      const pText = getParagraphPlainText(pEl);
      if (hasAnyKnownToken(pText, allTokenKeys)) {
        const replaced = replaceTokensInParagraph(xmlDoc, pEl, allTokenMap, templateCtx);
        if (replaced.size > 0) {
          replacedTokensMap.set(pEl, replaced);
        }
      }
    }

    if (xp === 'word/document.xml') {
      applyStructuralPopulations(xmlDoc, memo, docType, replacedTokensMap, stageMemo, templateCtx);
    } else {
      applyHeaderFooterMetadata(xmlDoc, memo, docType, replacedTokensMap, stageMemo);
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
        /^(?:Chronologic\s+Order\s+of\s+(?:the\s+)?IC\s*Memo|IC\s*Memo\s*(?:Number|No\.?))\b/i,
        ['#icMemoNumber#', '{{icMemoNumber}}', '#icMemoNo#']
      );
      checkNarrative(
        'IC Memo Date',
        si.icMemoDate,
        /^IC\s*Memo\s*Date\b/i,
        ['#icMemoDate#', '{{icMemoDate}}']
      );
      checkNarrative(
        'Recommendation of the Investment Committee (IC)',
        ss.mecIcRecommendation,
        /^(?:\d+\s*[.)]\s*)?(?:Recommendation\s+of\s+(?:the\s+)?Investment\s+Committee(?:\s*\(IC\))?|Investment\s+Committee\s*(?:\(IC\))?\s+Recommendation)(?:\s*\(Carried\s+Forward\))?\s*:?/i,
        ['#mecIcRecommendation#', '{{mecIcRecommendation}}']
      );
    } else if (docType === 'board') {
      checkNarrative(
        'Chronologic Order of the MEC Memo',
        si.mecMemoNumber,
        /^(?:Chronologic\s+Order\s+of\s+(?:the\s+)?MEC\s*Memo|MEC\s*Memo\s*(?:Number|No\.?))\b/i,
        ['#mecMemoNumber#', '{{mecMemoNumber}}', '#mecMemoNo#']
      );
      checkNarrative(
        'MEC Memo Date',
        si.mecMemoDate,
        /^MEC\s*Memo\s*Date\b/i,
        ['#mecMemoDate#', '{{mecMemoDate}}']
      );
      checkNarrative(
        'Management Executive Committee (MEC) Recommendation',
        ss.boardMecRecommendation,
        /^(?:\d+\s*[.)]\s*)?(?:Recommendation\s+of\s+(?:the\s+)?Management\s+Executive\s+Committee(?:\s*\(MEC\))?|Management\s+Executive\s+Committee\s*(?:\(MEC\))?\s+Recommendation)(?:\s*\(Carried\s+Forward\))?\s*:?/i,
        ['#boardMecRecommendation#', '{{boardMecRecommendation}}']
      );
    } else if (docType === 'sanction') {
      checkNarrative(
        'Chronologic Order of the Board Memo',
        si.boardMemoNumber,
        /^(?:Chronologic\s+Order\s+of\s+(?:the\s+)?Board\s*Memo|Board\s*Memo\s*(?:Number|No\.?))\b/i,
        ['#boardMemoNumber#', '{{boardMemoNumber}}', '#boardMemoNo#']
      );
      checkNarrative(
        'Board Memo Date',
        si.boardMemoDate,
        /^Board\s*Memo\s*Date\b/i,
        ['#boardMemoDate#', '{{boardMemoDate}}']
      );
      checkNarrative(
        'AI Cross-Verified Facility & Liability Confirmation',
        ss.sanctionCrossVerifiedSummary,
        /^(?:\d+\s*[.)]\s*)?AI\s+Cross-Verified\s+Facility(?:\s*(?:&|and)\s*Liability\s+Confirmation)?\s*:?/i,
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
