import type {
  AccountStatementData,
  BranchProposalData,
  BranchProposalSource,
  CibReportData,
  CostOfFundData,
  DocumentKey,
  OutstandingData,
  OverdueData,
  PartyPerformanceData,
} from '../types/memo';
import { cleanText, parseIndianNumber } from '../utils/numberFormat';
import {
  parseBranchProposalFromPdf,
  parseBranchProposalFromStructured,
} from './branchProposalParser';
import {
  extractPdfLines,
  PdfLine,
  PdfTextItem,
  renderPdfPagesToImages,
} from './pdfUtils';
import {
  extractDocxContent,
  extractExcelContent,
  isDocxFile,
  isExcelFile,
} from './structuredDocUtils';

const parseCache = new Map<string, unknown>();

function getFileSignature(prefix: string, file: File): string {
  return `${prefix}:${file.name}:${file.size}:${file.lastModified}`;
}

export function getCachedParse<T>(prefix: string, file: File): T | undefined {
  return parseCache.get(getFileSignature(prefix, file)) as T | undefined;
}

export function setCachedParse<T>(prefix: string, file: File, value: T): T {
  parseCache.set(getFileSignature(prefix, file), value);
  return value;
}

function findNearestItemByX(items: PdfTextItem[], targetX: number, maxDist = 65): PdfTextItem | undefined {
  let best: PdfTextItem | undefined;
  let bestDist = Infinity;
  for (const item of items) {
    const center = item.x + item.width / 2;
    const dist = Math.abs(center - targetX);
    if (dist < bestDist && dist <= maxDist) {
      bestDist = dist;
      best = item;
    }
  }
  return best;
}

/** Collect all 2D row grids (each table in DOCX or each sheet in Excel) for unified tabular scanning */
async function extractStructuredGrids(
  file: File,
  onStage?: StageCallback
): Promise<{
  grids: Array<Array<Array<string | number | null>>>;
  textGrids: string[][][];
  lines: string[];
  rawText: string;
}> {
  if (isDocxFile(file)) {
    onStage?.('Reading Word (.docx)...');
    const docx = await extractDocxContent(file);
    const textGrids = docx.tables.map((t) => t.rows);
    return {
      grids: textGrids,
      textGrids,
      lines: docx.lines,
      rawText: docx.rawText,
    };
  }

  onStage?.('Reading Excel workbook...');
  const excel = await extractExcelContent(file);
  return {
    grids: excel.sheets.map((s) => s.typedRows),
    textGrids: excel.sheets.map((s) => s.rows),
    lines: excel.lines,
    rawText: excel.rawText,
  };
}

function rowToText(row: Array<string | number | null>): string {
  return row
    .map((c) => (c === null || c === undefined ? '' : String(c).trim()))
    .filter(Boolean)
    .join(' ');
}

function extractDateFromText(text: string): string {
  const m =
    text.match(/(\d{1,2}-[A-Za-z]{3}-\d{2,4})/) ||
    text.match(/(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})/) ||
    text.match(/(\d{4}-\d{1,2}-\d{1,2})/) ||
    text.match(/(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{2,4})/);
  return m ? m[1] : '';
}

export type StageCallback = (stage: string) => void;

/**
 * Local non-AI OCR fallback using Tesseract.js — invoked ONLY for image files or
 * image-only PDFs that contain no embedded text layer.
 */
async function ocrImagesToPdfLines(
  images: Array<{ dataUrl: string; width: number; height: number }>,
  onStage?: StageCallback
): Promise<{ lines: PdfLine[]; rawText: string }> {
  const { createWorker } = await import('tesseract.js');
  const worker = await createWorker('eng');
  const allLines: PdfLine[] = [];
  const rawChunks: string[] = [];

  try {
    for (let i = 0; i < images.length; i++) {
      onStage?.(`Running non-AI OCR (page ${i + 1}/${images.length})...`);
      const img = images[i];
      const result = await worker.recognize(img.dataUrl, undefined, {
        blocks: true,
        text: true,
      });
      const scaleX = img.width > 0 ? 595.32 / img.width : 1;
      const scaleY = img.height > 0 ? 841.92 / img.height : 1;
      rawChunks.push(result.data.text || '');

      const pageItems: PdfTextItem[] = [];
      for (const block of result.data.blocks ?? []) {
        for (const para of block.paragraphs ?? []) {
          for (const tLine of para.lines ?? []) {
            for (const word of tLine.words ?? []) {
              const txt = word.text?.trim();
              if (!txt) continue;
              const x = word.bbox.x0 * scaleX;
              const y = 841.92 - word.bbox.y0 * scaleY;
              const width = (word.bbox.x1 - word.bbox.x0) * scaleX;
              const height = (word.bbox.y1 - word.bbox.y0) * scaleY;
              pageItems.push({ text: txt, x, y, width, height, page: i + 1 });
            }
          }
        }
      }

      pageItems.sort((a, b) => (Math.abs(b.y - a.y) > 4 ? b.y - a.y : a.x - b.x));
      const pageLines: PdfLine[] = [];
      for (const it of pageItems) {
        const existing = pageLines.find((l) => Math.abs(l.y - it.y) <= 4.5);
        if (existing) existing.items.push(it);
        else pageLines.push({ y: it.y, page: i + 1, items: [it], text: '' });
      }
      for (const l of pageLines) {
        l.items.sort((a, b) => a.x - b.x);
        l.text = l.items.map((it) => it.text).join(' ');
      }
      allLines.push(...pageLines);
    }
  } finally {
    await worker.terminate();
  }

  return { lines: allLines, rawText: rawChunks.join('\n') };
}

function getImageDimensions(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth || 1400, height: img.naturalHeight || 1980 });
    img.onerror = () => resolve({ width: 1400, height: 1980 });
    img.src = dataUrl;
  });
}

/**
 * 1. Branch Proposal (Ultra-fast Non-AI rule-based extraction across PDF, DOCX, Excel, and Image)
 */
export async function parseBranchProposalFile(
  file: File,
  onStage?: StageCallback
): Promise<BranchProposalData> {
  onStage?.('Checking cache...');
  const cached = getCachedParse<BranchProposalData>('bp', file);
  if (cached) {
    onStage?.('Loaded from cache');
    return cached;
  }

  if (isDocxFile(file)) {
    onStage?.('Reading Word (.docx)...');
    const docx = await extractDocxContent(file);
    if (!docx.rawText.trim()) {
      throw new Error('Could not extract any text or tables from the Word Branch Proposal.');
    }
    onStage?.('Extracting proposal fields...');
    const grids = docx.tables.map((t) => t.rows);
    const parsed = parseBranchProposalFromStructured(grids, docx.lines, docx.rawText);
    setCachedParse('bp_source', file, {
      sourceId: getFileSignature('src', file),
      fileName: file.name,
      rawText: docx.rawText,
      pageImages: [],
    } satisfies BranchProposalSource);
    return setCachedParse('bp', file, parsed);
  }

  if (isExcelFile(file)) {
    onStage?.('Reading Excel workbook...');
    const excel = await extractExcelContent(file);
    if (!excel.rawText.trim()) {
      throw new Error('Could not extract any worksheet content from the Excel Branch Proposal.');
    }
    onStage?.('Extracting proposal fields...');
    const grids = excel.sheets.map((s) => s.typedRows);
    const parsed = parseBranchProposalFromStructured(grids, excel.lines, excel.rawText);
    setCachedParse('bp_source', file, {
      sourceId: getFileSignature('src', file),
      fileName: file.name,
      rawText: excel.rawText,
      pageImages: [],
    } satisfies BranchProposalSource);
    return setCachedParse('bp', file, parsed);
  }

  if (file.type.startsWith('image/')) {
    onStage?.('Reading image...');
    const dataUrl = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result ?? ''));
      reader.onerror = () => reject(new Error('Failed to read image file.'));
      reader.readAsDataURL(file);
    });
    const dims = await getImageDimensions(dataUrl);
    const { lines, rawText } = await ocrImagesToPdfLines(
      [{ dataUrl, width: dims.width, height: dims.height }],
      onStage
    );
    onStage?.('Extracting proposal fields...');
    const parsed = parseBranchProposalFromPdf(lines, rawText);
    const base64 = dataUrl.split(',')[1] ?? '';
    setCachedParse('bp_source', file, {
      sourceId: getFileSignature('src', file),
      fileName: file.name,
      rawText,
      pageImages: base64 ? [{ mimeType: file.type || 'image/jpeg', data: base64 }] : [],
    } satisfies BranchProposalSource);
    return setCachedParse('bp', file, parsed);
  }

  // PDF: Read text layer directly at full speed; fall back to local non-AI Tesseract OCR only if image-only
  onStage?.('Reading Branch Proposal PDF...');
  const { lines, rawText } = await extractPdfLines(file);
  if (rawText.trim().length >= 80) {
    onStage?.('Extracting tables & fields...');
    const parsed = parseBranchProposalFromPdf(lines, rawText);
    setCachedParse('bp_raw', file, rawText);
    return setCachedParse('bp', file, parsed);
  }

  onStage?.('Rendering image-only PDF pages...');
  const rendered = await renderPdfPagesToImages(file, 1400, 10);
  if (rendered.length === 0) {
    throw new Error('Could not read text or render pages from the Branch Proposal PDF.');
  }
  const ocrResult = await ocrImagesToPdfLines(
    rendered.map((r) => ({
      dataUrl: `data:${r.mimeType};base64,${r.data}`,
      width: r.width,
      height: r.height,
    })),
    onStage
  );
  onStage?.('Extracting tables & fields...');
  const parsed = parseBranchProposalFromPdf(ocrResult.lines, ocrResult.rawText);
    setCachedParse('bp_source', file, {
      sourceId: getFileSignature('src', file),
      fileName: file.name,
      rawText: ocrResult.rawText,
      pageImages: rendered.map((r) => ({ mimeType: r.mimeType, data: r.data })),
    } satisfies BranchProposalSource);
  return setCachedParse('bp', file, parsed);
}

const inFlightSourcePrep = new Map<string, Promise<BranchProposalSource>>();

/**
 * Prepares the full Branch Proposal source (raw text + rendered page images for PDFs/images)
 * with promise deduplication & stable sourceId for instant repeated/parallel AI scrutiny.
 */
export async function prepareBranchProposalSource(file: File): Promise<BranchProposalSource> {
  const cached = getCachedParse<BranchProposalSource>('bp_source', file);
  if (cached) return cached;

  const sig = getFileSignature('bp_source', file);
  const existingPromise = inFlightSourcePrep.get(sig);
  if (existingPromise) return existingPromise;

  const sourceId = getFileSignature('src', file);

  const task = (async (): Promise<BranchProposalSource> => {
    if (isDocxFile(file)) {
      const docx = await extractDocxContent(file);
      return setCachedParse('bp_source', file, {
        sourceId,
        fileName: file.name,
        rawText: docx.rawText,
        pageImages: [],
      });
    }

    if (isExcelFile(file)) {
      const excel = await extractExcelContent(file);
      return setCachedParse('bp_source', file, {
        sourceId,
        fileName: file.name,
        rawText: excel.rawText,
        pageImages: [],
      });
    }

    if (file.type.startsWith('image/')) {
      const dataUrl = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result ?? ''));
        reader.onerror = () => reject(new Error('Failed to read image file.'));
        reader.readAsDataURL(file);
      });
      const base64 = dataUrl.split(',')[1] ?? '';
      const rawText = getCachedParse<string>('bp_raw', file) ?? '';
      return setCachedParse('bp_source', file, {
        sourceId,
        fileName: file.name,
        rawText,
        pageImages: base64 ? [{ mimeType: file.type || 'image/jpeg', data: base64 }] : [],
      });
    }

    // PDF: reuse cached rawText if already parsed and render JPEG page images in parallel
    let rawText = getCachedParse<string>('bp_raw', file);
    const [extractedText, rendered] = await Promise.all([
      rawText === undefined ? extractPdfLines(file).then((r) => r.rawText) : Promise.resolve(rawText),
      renderPdfPagesToImages(file, 840, 12),
    ]);
    if (rawText === undefined) {
      rawText = extractedText;
      setCachedParse('bp_raw', file, rawText);
    }

    return setCachedParse('bp_source', file, {
      sourceId,
      fileName: file.name,
      rawText,
      pageImages: rendered.map((r) => ({ mimeType: r.mimeType, data: r.data })),
    });
  })();

  inFlightSourcePrep.set(sig, task);
  try {
    return await task;
  } finally {
    inFlightSourcePrep.delete(sig);
  }
}

/**
 * Warms Branch Proposal multimodal page images asynchronously right after upload so
 * AI section generation has near-zero document preparation latency.
 */
export function warmBranchProposalSource(file: File): void {
  const trigger = () => {
    void prepareBranchProposalSource(file).catch(() => {
      // Ignore background warmup failures; on-demand call handles errors cleanly
    });
  };
  if (typeof window !== 'undefined' && 'requestIdleCallback' in window) {
    window.requestIdleCallback(trigger);
  } else {
    setTimeout(trigger, 0);
  }
}

/**
 * Prepares any completed stage memo upload (IC Memo, MEC Memo, or Board Memo)
 * in PDF, Word (.docx), Excel (.xlsx/.xls), or image format for multimodal AI extraction & scrutiny.
 */
export async function prepareMemoDocumentSource(
  file: File,
  onStage?: StageCallback
): Promise<BranchProposalSource> {
  const cached = getCachedParse<BranchProposalSource>('stage_memo_source', file);
  if (cached) {
    onStage?.('Loaded parsed document from cache...');
    return cached;
  }

  const sourceId = getFileSignature('stage_src', file);

  if (isDocxFile(file)) {
    onStage?.('Extracting Word (.docx) paragraphs and tables...');
    const docx = await extractDocxContent(file);
    return setCachedParse('stage_memo_source', file, {
      sourceId,
      fileName: file.name,
      rawText: docx.rawText,
      pageImages: [],
    });
  }

  if (isExcelFile(file)) {
    onStage?.('Extracting spreadsheet worksheets...');
    const excel = await extractExcelContent(file);
    return setCachedParse('stage_memo_source', file, {
      sourceId,
      fileName: file.name,
      rawText: excel.rawText,
      pageImages: [],
    });
  }

  if (file.type.startsWith('image/')) {
    onStage?.('Preparing image for AI visual scrutiny...');
    const dataUrl = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result ?? ''));
      reader.onerror = () => reject(new Error('Failed to read image file.'));
      reader.readAsDataURL(file);
    });
    const base64 = dataUrl.split(',')[1] ?? '';
    return setCachedParse('stage_memo_source', file, {
      sourceId,
      fileName: file.name,
      rawText: '',
      pageImages: base64 ? [{ mimeType: file.type || 'image/jpeg', data: base64 }] : [],
    });
  }

  onStage?.('Reading PDF pages and rendering page visuals...');
  const [extracted, rendered] = await Promise.all([
    extractPdfLines(file),
    renderPdfPagesToImages(file, 860, 14),
  ]);

  let rawText = extracted.rawText;
  if (rawText.trim().length < 250 && rendered.length > 0) {
    onStage?.('Running OCR fallback on scanned PDF pages...');
    try {
      const ocr = await ocrImagesToPdfLines(
        rendered.slice(0, 8).map((r) => ({
          dataUrl: `data:${r.mimeType};base64,${r.data}`,
          width: r.width,
          height: r.height,
        })),
        onStage
      );
      if (ocr.rawText.trim().length > rawText.trim().length) {
        rawText = ocr.rawText;
      }
    } catch {
      // Proceed with visual pageImages for multimodal Gemini scrutiny
    }
  }

  return setCachedParse('stage_memo_source', file, {
    sourceId,
    fileName: file.name,
    rawText,
    pageImages: rendered.map((r) => ({ mimeType: r.mimeType, data: r.data })),
  });
}

/**
 * 2. Outstanding (PDF, Word .docx, or Excel .xlsx/.xls)
 * Extracts Grand Total row: Princ Out, Charge, Profit Out, Compensation, Unapplied Pro, NET Out
 */
export async function parseOutstandingPdf(
  file: File,
  onStage?: StageCallback
): Promise<{
  data: OutstandingData;
  rawText: string;
}> {
  const cached = getCachedParse<{ data: OutstandingData; rawText: string }>('outstanding', file);
  if (cached) {
    onStage?.('Loaded from cache');
    return cached;
  }

  if (isDocxFile(file) || isExcelFile(file)) {
    const { grids, textGrids, rawText } = await extractStructuredGrids(file, onStage);
    onStage?.('Extracting balances...');

    let princOutBDT: number | null = null;
    let chargeBDT: number | null = null;
    let profitOutBDT: number | null = null;
    let compensationBDT: number | null = null;
    let unappliedProBDT: number | null = null;
    let netOutBDT: number | null = null;

    for (let g = 0; g < grids.length; g++) {
      const grid = grids[g];
      const textGrid = textGrids[g];

      const headerIdx = textGrid.findIndex((r) => {
        const line = r.join(' ');
        return /princ\.?\s*out|principal/i.test(line) && /net\s*out/i.test(line);
      });

      const grandTotalIdx = textGrid.findIndex((r) => /grand\s*total/i.test(r.join(' ')));
      const fallbackTotalIdx =
        grandTotalIdx !== -1
          ? grandTotalIdx
          : textGrid.findIndex((r) => /\btotal\b/i.test(r.join(' ')));

      if (headerIdx !== -1 && fallbackTotalIdx !== -1) {
        const headerRow = textGrid[headerIdx];
        const nextRow = textGrid[headerIdx + 1] ?? [];
        const combinedHeaders = headerRow.map((h, i) => `${h ?? ''} ${nextRow[i] ?? ''}`.trim());

        const findCol = (pattern: RegExp): number =>
          combinedHeaders.findIndex((cellText) => pattern.test(cellText));

        const princCol = findCol(/princ\.?\s*out|principal/i);
        const profitCol = findCol(/profit\s*out/i);
        const compCol = findCol(/compensation/i);
        const unappCol = findCol(/unapplied/i);
        const chargeCol = findCol(/charge/i);
        const netCol = findCol(/net\s*out/i);

        const totalRow = grid[fallbackTotalIdx];
        if (princCol >= 0) princOutBDT = parseIndianNumber(totalRow[princCol]);
        if (profitCol >= 0) profitOutBDT = parseIndianNumber(totalRow[profitCol]);
        if (compCol >= 0) compensationBDT = parseIndianNumber(totalRow[compCol]);
        if (unappCol >= 0) unappliedProBDT = parseIndianNumber(totalRow[unappCol]);
        if (chargeCol >= 0) chargeBDT = parseIndianNumber(totalRow[chargeCol]);
        if (netCol >= 0) netOutBDT = parseIndianNumber(totalRow[netCol]);

        // Positional fallback if merged headers misaligned
        if (netOutBDT === null) {
          const nums = totalRow
            .map((c) => parseIndianNumber(c))
            .filter((n): n is number => n !== null);
          if (nums.length >= 8) {
            princOutBDT = princOutBDT ?? nums[1];
            profitOutBDT = profitOutBDT ?? nums[2];
            compensationBDT = compensationBDT ?? nums[3];
            unappliedProBDT = unappliedProBDT ?? nums[4];
            chargeBDT = chargeBDT ?? nums[5];
            netOutBDT = netOutBDT ?? nums[7];
          }
        }
        if (netOutBDT !== null || princOutBDT !== null) break;
      }
    }

    return setCachedParse('outstanding', file, {
      data: {
        princOutBDT,
        chargeBDT,
        profitOutBDT,
        compensationBDT,
        unappliedProBDT,
        netOutBDT,
      },
      rawText,
    });
  }

  onStage?.('Reading PDF text...');
  const { lines, rawText } = await extractPdfLines(file);
  onStage?.('Extracting balances...');

  const headerLine = lines.find(
    (l) => /princ\s*out/i.test(l.text) && /net\s*out/i.test(l.text)
  );
  const grandTotalLine = lines.find((l) => /grand\s*total/i.test(l.text));

  let princOutBDT: number | null = null;
  let chargeBDT: number | null = null;
  let profitOutBDT: number | null = null;
  let compensationBDT: number | null = null;
  let unappliedProBDT: number | null = null;
  let netOutBDT: number | null = null;

  if (headerLine && grandTotalLine) {
    const findColX = (pattern: RegExp): number | null => {
      const match = headerLine.items.find((it) => pattern.test(it.text));
      return match ? match.x + match.width / 2 : null;
    };

    const princX = findColX(/princ\s*out/i);
    const profitX = findColX(/profit\s*out/i);
    const compX = findColX(/compensation/i);
    const unappX = findColX(/unapplied/i);
    const chargeX = findColX(/charge/i);
    const netX = findColX(/net\s*out/i);

    const numericItems = grandTotalLine.items.filter(
      (it) => parseIndianNumber(it.text) !== null
    );

    if (princX !== null) {
      const m = findNearestItemByX(numericItems, princX);
      if (m) princOutBDT = parseIndianNumber(m.text);
    }
    if (profitX !== null) {
      const m = findNearestItemByX(numericItems, profitX);
      if (m) profitOutBDT = parseIndianNumber(m.text);
    }
    if (compX !== null) {
      const m = findNearestItemByX(numericItems, compX);
      if (m) compensationBDT = parseIndianNumber(m.text);
    }
    if (unappX !== null) {
      const m = findNearestItemByX(numericItems, unappX);
      if (m) unappliedProBDT = parseIndianNumber(m.text);
    }
    if (chargeX !== null) {
      const m = findNearestItemByX(numericItems, chargeX);
      if (m) chargeBDT = parseIndianNumber(m.text);
    }
    if (netX !== null) {
      const m = findNearestItemByX(numericItems, netX);
      if (m) netOutBDT = parseIndianNumber(m.text);
    }
  }

  if (grandTotalLine && netOutBDT === null) {
    const nums = grandTotalLine.text
      .replace(/grand\s*total\s*:?/i, '')
      .trim()
      .split(/\s+/)
      .map((tok) => parseIndianNumber(tok))
      .filter((n): n is number => n !== null);
    if (nums.length >= 8) {
      princOutBDT = princOutBDT ?? nums[1];
      profitOutBDT = profitOutBDT ?? nums[2];
      compensationBDT = compensationBDT ?? nums[3];
      unappliedProBDT = unappliedProBDT ?? nums[4];
      chargeBDT = chargeBDT ?? nums[5];
      netOutBDT = netOutBDT ?? nums[7];
    }
  }

  return setCachedParse('outstanding', file, {
    data: {
      princOutBDT,
      chargeBDT,
      profitOutBDT,
      compensationBDT,
      unappliedProBDT,
      netOutBDT,
    },
    rawText,
  });
}

/**
 * 3. Party Performance (PDF, Word .docx, or Excel .xlsx/.xls)
 * Extracts Product Name (#5#/#55#), Customer Total Recovery Amt (#60#), Print Date (#68#)
 */
export async function parsePartyPerformancePdf(
  file: File,
  onStage?: StageCallback
): Promise<{
  data: PartyPerformanceData;
  rawText: string;
}> {
  const cached = getCachedParse<{ data: PartyPerformanceData; rawText: string }>('partyPerf', file);
  if (cached) {
    onStage?.('Loaded from cache');
    return cached;
  }

  if (isDocxFile(file) || isExcelFile(file)) {
    const { grids, textGrids, lines, rawText } = await extractStructuredGrids(file, onStage);
    onStage?.('Extracting performance figures...');

    let productName = '';
    let printDate = '';
    let recoveryAmtBDT: number | null = null;

    for (const line of lines) {
      if (!printDate) {
        const m = line.match(/Print\s*Date\s*:?\s*([0-9./-]+)/i);
        if (m) printDate = cleanText(m[1]);
      }
      if (!productName) {
        const m = line.match(/Product\s*Name\s*:?\s*(.+)/i);
        if (m) {
          // Strip trailing pipe-separated metadata if any
          productName = cleanText(m[1].split('|')[0]);
        }
      }
    }

    for (let g = 0; g < grids.length; g++) {
      const grid = grids[g];
      const textGrid = textGrids[g];

      for (let r = 0; r < textGrid.length; r++) {
        const row = textGrid[r];
        for (let c = 0; c < row.length; c++) {
          const cellTxt = row[c] ?? '';
          if (!printDate && /Print\s*Date/i.test(cellTxt)) {
            const inline = extractDateFromText(cellTxt);
            if (inline) {
              printDate = inline;
            } else {
              for (let k = c + 1; k < row.length; k++) {
                if (row[k]?.trim()) {
                  printDate = extractDateFromText(row[k]) || cleanText(row[k]);
                  break;
                }
              }
            }
          }
          if (!productName && /Product\s*Name/i.test(cellTxt)) {
            const m = cellTxt.match(/Product\s*Name\s*:?\s*(.+)/i);
            if (m && cleanText(m[1])) {
              productName = cleanText(m[1]);
            } else {
              for (let k = c + 1; k < row.length; k++) {
                if (row[k]?.trim()) {
                  productName = cleanText(row[k]);
                  break;
                }
              }
            }
          }
        }
      }

      // Find Recovery Amt column index
      let recoveryColIdx = -1;
      const recHeaderIdx = textGrid.findIndex((r) => /recovery\s*amt/i.test(r.join(' ')));
      if (recHeaderIdx !== -1) {
        recoveryColIdx = textGrid[recHeaderIdx].findIndex((c) => /recovery\s*amt/i.test(c));
      }

      // Locate Customer Total row / block
      const custTotalIdx = textGrid.findIndex((r) => /customer\s*total/i.test(r.join(' ')));
      if (custTotalIdx !== -1) {
        const windowEnd = Math.min(textGrid.length, custTotalIdx + 5);
        const windowStart = Math.max(0, custTotalIdx - 2);
        const candidates: number[] = [];

        for (let r = windowStart; r < windowEnd; r++) {
          const rowText = textGrid[r].join(' ');
          if (/product\s*total|prepared\s*by/i.test(rowText)) continue;

          // If stacked block has an explicit "Total" line below Customer Total, prefer that row
          const isExplicitBottomTotal = /\bTotal\b/i.test(rowText) && !/customer\s*total/i.test(rowText);
          let val: number | null = null;
          if (recoveryColIdx >= 0) {
            val = parseIndianNumber(grid[r][recoveryColIdx]);
          } else {
            const nums = grid[r]
              .map((cellVal) => parseIndianNumber(cellVal))
              .filter((n): n is number => n !== null);
            if (nums.length >= 2) val = nums[1];
            else if (nums.length === 1) val = nums[0];
          }
          if (val !== null) {
            if (isExplicitBottomTotal) {
              recoveryAmtBDT = val;
              break;
            }
            candidates.push(val);
          }
        }

        if (recoveryAmtBDT === null && candidates.length > 0) {
          recoveryAmtBDT = candidates[candidates.length - 1];
        }
      }

      if (recoveryAmtBDT !== null) break;
    }

    return setCachedParse('partyPerf', file, {
      data: {
        productName,
        recoveryAmtBDT,
        printDate,
      },
      rawText,
    });
  }

  onStage?.('Reading PDF text...');
  const { lines, rawText } = await extractPdfLines(file);
  onStage?.('Extracting performance figures...');

  let productName = '';
  let printDate = '';
  let recoveryAmtBDT: number | null = null;

  for (const line of lines) {
    const m = line.text.match(/Print\s*Date\s*:\s*([0-9./-]+)/i);
    if (m) {
      printDate = cleanText(m[1]);
      break;
    }
  }

  for (const line of lines) {
    const m = line.text.match(/Product\s*Name\s*:?\s*(.+)/i);
    if (m) {
      productName = cleanText(m[1]);
      break;
    }
  }

  const recHeaderLine = lines.find((l) => /recovery\s*amt/i.test(l.text));
  let recoveryColX: number | null = null;
  if (recHeaderLine) {
    const recItem = recHeaderLine.items.find((it) => /recovery/i.test(it.text));
    if (recItem) {
      recoveryColX = recItem.x + recItem.width / 2;
    }
  }

  const custTotalIdx = lines.findIndex((l) => /customer\s*total/i.test(l.text));
  if (custTotalIdx !== -1) {
    const windowLines = lines
      .slice(Math.max(0, custTotalIdx - 2), custTotalIdx + 4)
      .filter((l) => !/product\s*total|prepared\s*by/i.test(l.text));
    const recoveryNumbers: Array<{ y: number; val: number }> = [];

    for (const wl of windowLines) {
      if (recoveryColX !== null) {
        const match = findNearestItemByX(
          wl.items.filter((it) => parseIndianNumber(it.text) !== null),
          recoveryColX,
          55
        );
        if (match) {
          const v = parseIndianNumber(match.text);
          if (v !== null) recoveryNumbers.push({ y: wl.y, val: v });
        }
      } else {
        const nums = wl.items
          .map((it) => parseIndianNumber(it.text))
          .filter((n): n is number => n !== null);
        if (nums.length >= 2) {
          recoveryNumbers.push({ y: wl.y, val: nums[1] });
        }
      }
    }

    if (recoveryNumbers.length > 0) {
      recoveryNumbers.sort((a, b) => a.y - b.y);
      recoveryAmtBDT = recoveryNumbers[0].val;
    }
  }

  return setCachedParse('partyPerf', file, {
    data: {
      productName,
      recoveryAmtBDT,
      printDate,
    },
    rawText,
  });
}

/**
 * 4. Overdue (PDF, Word .docx, or Excel .xlsx/.xls)
 * Extracts Grand Total of Overdue Amount -> #59#
 */
export async function parseOverduePdf(
  file: File,
  onStage?: StageCallback
): Promise<{
  data: OverdueData;
  rawText: string;
}> {
  const cached = getCachedParse<{ data: OverdueData; rawText: string }>('overdue', file);
  if (cached) {
    onStage?.('Loaded from cache');
    return cached;
  }

  if (isDocxFile(file) || isExcelFile(file)) {
    const { grids, textGrids, rawText } = await extractStructuredGrids(file, onStage);
    onStage?.('Extracting overdue total...');

    let overdueAmountBDT: number | null = null;
    for (let g = 0; g < grids.length; g++) {
      const grid = grids[g];
      const textGrid = textGrids[g];

      const headerIdx = textGrid.findIndex((r) => /overdue/i.test(r.join(' ')));
      let overdueCol = -1;
      if (headerIdx !== -1) {
        // Prefer specific "Overdue Amount" / "Overdue Amt" column if present
        overdueCol = textGrid[headerIdx].findIndex((c) => /overdue\s*(amt|amount)/i.test(c));
        if (overdueCol === -1) {
          overdueCol = textGrid[headerIdx].findIndex((c) => /\boverdue\b/i.test(c));
        }
      }

      const grandTotalIdx = textGrid.findIndex((r) => /grand\s*total/i.test(r.join(' ')));
      const totalIdx =
        grandTotalIdx !== -1
          ? grandTotalIdx
          : textGrid.findIndex((r) => /\btotal\b/i.test(r.join(' ')));

      if (totalIdx !== -1) {
        const totalRow = grid[totalIdx];
        if (overdueCol >= 0) {
          const val = parseIndianNumber(totalRow[overdueCol]);
          if (val !== null) {
            overdueAmountBDT = val;
            break;
          }
        }
        const nums = totalRow
          .map((c) => parseIndianNumber(c))
          .filter((n): n is number => n !== null);
        if (nums.length > 0) {
          overdueAmountBDT = nums[nums.length - 1];
          break;
        }
      }
    }

    return setCachedParse('overdue', file, {
      data: { overdueAmountBDT },
      rawText,
    });
  }

  onStage?.('Reading PDF text...');
  const { lines, rawText } = await extractPdfLines(file);
  onStage?.('Extracting overdue total...');

  let overdueAmountBDT: number | null = null;
  const grandTotalLine = lines.find((l) => /grand\s*total/i.test(l.text));

  if (grandTotalLine) {
    const nums = grandTotalLine.items
      .map((it) => parseIndianNumber(it.text))
      .filter((n): n is number => n !== null);
    if (nums.length > 0) {
      overdueAmountBDT = nums[nums.length - 1];
    }
  }

  return setCachedParse('overdue', file, {
    data: { overdueAmountBDT },
    rawText,
  });
}

/**
 * 5. Account Statement (PDF, Word .docx, or Excel .xlsx/.xls)
 * Extracts #2# Account number, #36# A/C No., #3# Branch name, #69# Customer ID, #35# Last row Balance
 */
export async function parseAccountStatementPdf(
  file: File,
  onStage?: StageCallback
): Promise<{
  data: AccountStatementData;
  rawText: string;
}> {
  const cached = getCachedParse<{ data: AccountStatementData; rawText: string }>('acctStmt', file);
  if (cached) {
    onStage?.('Loaded from cache');
    return cached;
  }

  if (isDocxFile(file) || isExcelFile(file)) {
    const { grids, textGrids, lines, rawText } = await extractStructuredGrids(file, onStage);
    onStage?.('Extracting account & balance...');

    let branchName = '';
    let customerId = '';
    let acNo = '';
    let lastBalanceBDT: number | null = null;

    // 1. Scan lines for explicit labels & 2nd-line branch fallback
    const cleanHeaderLines = lines
      .slice(0, 12)
      .map((l) => cleanText(l.replace(/print\s*date\s*:?.*$/i, '').replace(/\|/g, ' ')))
      .filter(Boolean);

    for (const line of lines) {
      if (!customerId) {
        const cidMatch = line.match(/Customer\s*ID\s*:?\s*([A-Za-z0-9-]+)/i);
        if (cidMatch) customerId = cleanText(cidMatch[1]);
      }
      if (!acNo) {
        const acMatch = line.match(/(?:A\/C|Account)\s*(?:No\.?|Number)\s*:?\s*([A-Za-z0-9-]+)/i);
        if (acMatch) acNo = cleanText(acMatch[1]);
      }
      if (!branchName) {
        const brMatch = line.match(/\bBranch(?:\s*Name)?\s*:?\s*([A-Za-z0-9 .,&()-]+)/i);
        if (brMatch && !/print\s*date/i.test(brMatch[1])) {
          branchName = cleanText(brMatch[1].split('|')[0]);
        }
      }
    }

    // Also scan grid cells for key-value adjacent pairs (e.g., ["Customer ID", "215675"])
    for (let g = 0; g < grids.length; g++) {
      const grid = grids[g];
      const textGrid = textGrids[g];
      for (let r = 0; r < textGrid.length; r++) {
        const row = textGrid[r];
        for (let c = 0; c < row.length; c++) {
          const cellTxt = row[c] ?? '';
          if (!customerId && /^Customer\s*ID\s*:?$/i.test(cellTxt.trim())) {
            for (let k = c + 1; k < row.length; k++) {
              if (row[k]?.trim()) {
                customerId = cleanText(row[k]);
                break;
              }
            }
          }
          if (!acNo && /^(?:A\/C|Account)\s*(?:No\.?|Number)\s*:?$/i.test(cellTxt.trim())) {
            for (let k = c + 1; k < row.length; k++) {
              if (row[k]?.trim()) {
                acNo = cleanText(row[k]);
                break;
              }
            }
          }
          if (!branchName && /^Branch(?:\s*Name)?\s*:?$/i.test(cellTxt.trim())) {
            for (let k = c + 1; k < row.length; k++) {
              if (row[k]?.trim()) {
                branchName = cleanText(row[k]);
                break;
              }
            }
          }
        }
      }

      // Transaction table: header with Debit, Credit, and Balance
      const headerIdx = textGrid.findIndex((r) => {
        const txt = r.join(' ');
        return /\bDebit\b/i.test(txt) && /\bCredit\b/i.test(txt) && /\bBalance\b/i.test(txt);
      });

      if (headerIdx !== -1 && lastBalanceBDT === null) {
        const balCol = textGrid[headerIdx].findIndex((c) => /\bBalance\b/i.test(c));
        const totalIdx = textGrid.findIndex(
          (r, idx) => idx > headerIdx && /\bTotal\s*:?/i.test(r.join(' '))
        );
        const endIdx = totalIdx > headerIdx ? totalIdx : textGrid.length;

        const candidateBalances: number[] = [];
        for (let r = headerIdx + 1; r < endIdx; r++) {
          const rowTxt = rowToText(grid[r]);
          const looksLikeTx =
            /^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}/.test(rowTxt) ||
            /^\d{4}-\d{1,2}-\d{1,2}/.test(rowTxt) ||
            /Balance\s*B\/F/i.test(rowTxt) ||
            (balCol >= 0 && parseIndianNumber(grid[r][balCol]) !== null);
          if (!looksLikeTx) continue;

          if (balCol >= 0) {
            const val = parseIndianNumber(grid[r][balCol]);
            if (val !== null) candidateBalances.push(val);
          } else {
            const nums = grid[r]
              .map((cellVal) => parseIndianNumber(cellVal))
              .filter((n): n is number => n !== null);
            if (nums.length > 0) candidateBalances.push(nums[nums.length - 1]);
          }
        }

        if (candidateBalances.length > 0) {
          lastBalanceBDT = candidateBalances[candidateBalances.length - 1];
        }
      }
    }

    // Fallback to 2nd line of statement header (matches PDF rule)
    if (!branchName && cleanHeaderLines.length >= 2) {
      const candidate = cleanHeaderLines[1];
      if (!/customer\s*id|a\/c|statement|debit|credit/i.test(candidate)) {
        branchName = candidate;
      }
    }

    return setCachedParse('acctStmt', file, {
      data: {
        accountNumber: acNo,
        acNo,
        branchName,
        customerId,
        lastBalanceBDT,
      },
      rawText,
    });
  }

  onStage?.('Reading PDF text...');
  const { lines, rawText } = await extractPdfLines(file);
  onStage?.('Extracting account & balance...');

  let branchName = '';
  const page1Lines = lines.filter((l) => l.page === 1);
  if (page1Lines.length >= 2) {
    const secondLineCleaned = page1Lines[1].items
      .filter((it) => !/print\s*date/i.test(it.text))
      .map((it) => it.text.trim())
      .join(' ');
    branchName = cleanText(secondLineCleaned);
  }

  let customerId = '';
  let acNo = '';

  for (const line of lines) {
    const cidMatch = line.text.match(/Customer\s*ID\s*:\s*([A-Za-z0-9-]+)/i);
    if (cidMatch && !customerId) {
      customerId = cleanText(cidMatch[1]);
    }
    const acMatch = line.text.match(/A\/C\s*No\.?\s*:\s*([A-Za-z0-9-]+)/i);
    if (acMatch && !acNo) {
      acNo = cleanText(acMatch[1]);
    }
  }

  let lastBalanceBDT: number | null = null;
  const headerLine = lines.find(
    (l) => /\bDebit\b/i.test(l.text) && /\bCredit\b/i.test(l.text) && /\bBalance\b/i.test(l.text)
  );
  const totalLineIdx = lines.findIndex((l) => /^Total\s*:/i.test(l.text) || /\bTotal\s*:/i.test(l.text));

  if (headerLine) {
    const balHeaderItem = headerLine.items.find((it) => /\bBalance\b/i.test(it.text));
    const balX = balHeaderItem ? balHeaderItem.x + balHeaderItem.width / 2 : null;
    const headerIdx = lines.indexOf(headerLine);
    const endIdx = totalLineIdx > headerIdx ? totalLineIdx : lines.length;

    const candidateBalances: number[] = [];
    for (let i = headerIdx + 1; i < endIdx; i++) {
      const row = lines[i];
      const isTxRow = /^\d{2}\/\d{2}\/\d{4}/.test(row.text) || /Balance\s*B\/F/i.test(row.text);
      if (!isTxRow) continue;

      if (balX !== null) {
        const numItems = row.items.filter((it) => parseIndianNumber(it.text) !== null);
        const balItem = findNearestItemByX(numItems, balX, 75);
        if (balItem) {
          const val = parseIndianNumber(balItem.text);
          if (val !== null) candidateBalances.push(val);
        }
      } else {
        const nums = row.items
          .map((it) => parseIndianNumber(it.text))
          .filter((n): n is number => n !== null);
        if (nums.length > 0) candidateBalances.push(nums[nums.length - 1]);
      }
    }

    if (candidateBalances.length > 0) {
      lastBalanceBDT = candidateBalances[candidateBalances.length - 1];
    }
  }

  return setCachedParse('acctStmt', file, {
    data: {
      accountNumber: acNo,
      acNo,
      branchName,
      customerId,
      lastBalanceBDT,
    },
    rawText,
  });
}

/**
 * 6. CIB Report (PDF, Word .docx, or Excel .xlsx/.xls)
 * Extracts #62# CIB date from Date of Inquiry
 */
export async function parseCibReportPdf(
  file: File,
  onStage?: StageCallback
): Promise<{
  data: CibReportData;
  rawText: string;
}> {
  const cached = getCachedParse<{ data: CibReportData; rawText: string }>('cib', file);
  if (cached) {
    onStage?.('Loaded from cache');
    return cached;
  }

  if (isDocxFile(file) || isExcelFile(file)) {
    const { textGrids, lines, rawText } = await extractStructuredGrids(file, onStage);
    onStage?.('Extracting inquiry date...');

    let cibDate = '';

    // Check grids (inline, right cell, or cell directly below "Date of Inquiry")
    for (const grid of textGrids) {
      for (let r = 0; r < grid.length; r++) {
        for (let c = 0; c < grid[r].length; c++) {
          if (/Date\s*of\s*Inquiry/i.test(grid[r][c] ?? '')) {
            const inline = extractDateFromText(grid[r][c]);
            if (inline) {
              cibDate = inline;
              break;
            }
            for (let k = c + 1; k < grid[r].length; k++) {
              const d = extractDateFromText(grid[r][k] ?? '');
              if (d) {
                cibDate = d;
                break;
              }
            }
            if (!cibDate && grid[r + 1]?.[c]) {
              const d = extractDateFromText(grid[r + 1][c]);
              if (d) cibDate = d;
            }
            if (!cibDate && grid[r + 1]) {
              const d = extractDateFromText(grid[r + 1].join(' '));
              if (d) cibDate = d;
            }
          }
          if (cibDate) break;
        }
        if (cibDate) break;
      }
      if (cibDate) break;
    }

    if (!cibDate) {
      const idx = lines.findIndex((l) => /Date\s*of\s*Inquiry/i.test(l));
      if (idx !== -1) {
        for (const l of lines.slice(idx, idx + 3)) {
          const d = extractDateFromText(l);
          if (d) {
            cibDate = d;
            break;
          }
        }
      }
    }

    if (!cibDate) {
      cibDate = extractDateFromText(lines.slice(0, 15).join(' '));
    }

    return setCachedParse('cib', file, {
      data: { cibDate },
      rawText,
    });
  }

  onStage?.('Reading PDF text...');
  const { lines, rawText } = await extractPdfLines(file);
  onStage?.('Extracting inquiry date...');

  let cibDate = '';

  const inquiryHeaderIdx = lines.findIndex((l) => /Date\s*of\s*Inquiry/i.test(l.text));
  if (inquiryHeaderIdx !== -1) {
    const searchWindow = lines.slice(inquiryHeaderIdx, inquiryHeaderIdx + 3);
    for (const l of searchWindow) {
      const m =
        l.text.match(/(\d{1,2}-[A-Za-z]{3}-\d{4})/) ||
        l.text.match(/(\d{1,2}[./-]\d{1,2}[./-]\d{4})/);
      if (m) {
        cibDate = m[1];
        break;
      }
    }
  }

  if (!cibDate) {
    const firstPage = lines.slice(0, 15).map((l) => l.text).join(' ');
    const m =
      firstPage.match(/(\d{1,2}-[A-Za-z]{3}-\d{4})/) ||
      firstPage.match(/(\d{1,2}\s+[A-Za-z]{3}\s+\d{4})/);
    if (m) cibDate = m[1];
  }

  return setCachedParse('cib', file, {
    data: { cibDate },
    rawText,
  });
}

/**
 * 7. Cost of Fund (.xlsx / .xls / .csv, Word .docx, or .pdf)
 * Extracts "Surplus/Deficit of realization over cost of fund" -> #31#
 */
export async function parseCostOfFundFile(
  file: File,
  onStage?: StageCallback
): Promise<{
  data: CostOfFundData;
  rawText: string;
}> {
  const cached = getCachedParse<{ data: CostOfFundData; rawText: string }>('cof', file);
  if (cached) {
    onStage?.('Loaded from cache');
    return cached;
  }

  if (isDocxFile(file)) {
    onStage?.('Reading Word (.docx)...');
    const docx = await extractDocxContent(file);
    onStage?.('Extracting cost of fund...');
    let surplusDeficitBDT: number | null = null;

    for (const tbl of docx.tables) {
      for (const row of tbl.rows) {
        const hasTarget = row.some((c) =>
          /surpluse?\s*\/?\s*deficit\s+of\s+realization/i.test(c)
        );
        if (hasTarget) {
          for (let c = row.length - 1; c >= 0; c--) {
            const val = parseIndianNumber(row[c]);
            if (val !== null) {
              surplusDeficitBDT = val;
              break;
            }
          }
          if (surplusDeficitBDT !== null) break;
        }
      }
      if (surplusDeficitBDT !== null) break;
    }

    if (surplusDeficitBDT === null) {
      for (const line of docx.lines) {
        if (/surpluse?\s*\/?\s*deficit\s+of\s+realization/i.test(line)) {
          const tokens = line.split(/\s+|\|/);
          for (let i = tokens.length - 1; i >= 0; i--) {
            const val = parseIndianNumber(tokens[i]);
            if (val !== null) {
              surplusDeficitBDT = val;
              break;
            }
          }
          if (surplusDeficitBDT !== null) break;
        }
      }
    }

    return setCachedParse('cof', file, { data: { surplusDeficitBDT }, rawText: docx.rawText });
  }

  if (isExcelFile(file)) {
    onStage?.('Reading Excel workbook...');
    const excel = await extractExcelContent(file);
    onStage?.('Scanning sheets...');

    let surplusDeficitBDT: number | null = null;

    for (const sheet of excel.sheets) {
      for (let r = 0; r < sheet.rows.length; r++) {
        const dispRow = sheet.rows[r];
        const typedRow = sheet.typedRows[r];
        const hasTargetLabel = dispRow.some((cell) =>
          /surpluse?\s*\/?\s*deficit\s+of\s+realization/i.test(cell)
        );

        if (hasTargetLabel) {
          for (let c = typedRow.length - 1; c >= 0; c--) {
            const val = parseIndianNumber(typedRow[c] ?? dispRow[c]);
            if (val !== null) {
              surplusDeficitBDT = val;
              break;
            }
          }
          if (surplusDeficitBDT !== null) break;
        }
      }
      if (surplusDeficitBDT !== null) break;
    }

    return setCachedParse('cof', file, {
      data: { surplusDeficitBDT },
      rawText: excel.rawText,
    });
  }

  // PDF fallback
  onStage?.('Reading PDF text...');
  const { lines, rawText } = await extractPdfLines(file);
  onStage?.('Extracting cost of fund...');
  let surplusDeficitBDT: number | null = null;
  for (const line of lines) {
    if (/surpluse?\s*\/?\s*deficit\s+of\s+realization\s+over\s+cost\s+of\s+fund/i.test(line.text)) {
      const nums = line.items
        .map((it) => parseIndianNumber(it.text))
        .filter((n): n is number => n !== null);
      if (nums.length > 0) {
        surplusDeficitBDT = nums[nums.length - 1];
        break;
      }
    }
  }
  return setCachedParse('cof', file, { data: { surplusDeficitBDT }, rawText });
}

function matchDocumentKeyFromText(rawText: string): DocumentKey | null {
  if (/customer\s+outstanding\s+report|princ\.?\s+out/i.test(rawText)) return 'outstanding';
  if (/perform[ea]nce\s+report|recovery\s+amt/i.test(rawText)) return 'partyPerformance';
  if (/statement\s+of\s+overdue/i.test(rawText)) return 'overdue';
  if (/a\/c\s+statement|mudaraba|balance\s+b\/f/i.test(rawText)) return 'accountStatement';
  if (/credit\s+information\s+bureau|cib\s+subject\s+code|date\s+of\s+inquiry/i.test(rawText)) return 'cibReport';
  if (/cost\s+of\s+fund|realization\s+over\s+cost/i.test(rawText)) return 'costOfFund';
  if (/exit\s+with\s+waiver|justification\s+of\s+the\s+proposed\s+facility|special\s+asset\s+management/i.test(rawText)) {
    return 'branchProposal';
  }
  return null;
}

/**
 * Automatically identifies which of the 7 banking documents a file corresponds to
 * across PDF, DOCX, Excel (.xlsx/.xls), and Image formats for instant batch upload.
 */
export async function detectDocumentKeyFromFile(file: File): Promise<DocumentKey | null> {
  const name = file.name.toLowerCase();

  if (/branch.*proposal|proposal/.test(name) || file.type.startsWith('image/')) {
    return 'branchProposal';
  }
  if (/outstanding/.test(name)) {
    return 'outstanding';
  }
  if (/party.*performance|performance/.test(name)) {
    return 'partyPerformance';
  }
  if (/overdue/.test(name)) {
    return 'overdue';
  }
  if (/account.*statement|statement|a\/c/.test(name)) {
    return 'accountStatement';
  }
  if (/cib/.test(name)) {
    return 'cibReport';
  }
  if (/cost.*fund|cof/.test(name)) {
    return 'costOfFund';
  }

  // Content inspection fallback for PDF, DOCX, and Excel
  try {
    if (isDocxFile(file)) {
      const { rawText } = await extractDocxContent(file);
      return matchDocumentKeyFromText(rawText);
    }
    if (isExcelFile(file)) {
      const { rawText } = await extractExcelContent(file);
      return matchDocumentKeyFromText(rawText) ?? 'costOfFund';
    }
    const { rawText } = await extractPdfLines(file);
    const matched = matchDocumentKeyFromText(rawText);
    if (matched) return matched;
    // Scanned PDF with negligible extractable text -> likely scanned Branch Proposal
    if (rawText.trim().length < 150) return 'branchProposal';
  } catch {
    // Ignore inspection error
  }

  return null;
}
