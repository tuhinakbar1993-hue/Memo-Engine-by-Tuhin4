import * as pdfjsLib from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

// Pre-warm a single PDFWorker instance so first upload has zero worker startup lag
let sharedWorker: pdfjsLib.PDFWorker | null = null;
function getSharedPdfWorker(): pdfjsLib.PDFWorker {
  if (!sharedWorker) {
    sharedWorker = new pdfjsLib.PDFWorker();
  }
  return sharedWorker;
}

// Pre-warm immediately in browser idle time
if (typeof window !== 'undefined') {
  const warm = () => {
    try {
      getSharedPdfWorker();
    } catch {
      // Ignore pre-warm errors
    }
  };
  if ('requestIdleCallback' in window) {
    window.requestIdleCallback(warm);
  } else {
    setTimeout(warm, 150);
  }
}

export interface PdfTextItem {
  text: string;
  x: number;
  y: number;
  width: number;
  height: number;
  page: number;
}

export interface PdfLine {
  y: number;
  page: number;
  items: PdfTextItem[];
  text: string;
}

const pdfLinesCache = new Map<
  string,
  { lines: PdfLine[]; rawText: string; pageCount: number }
>();

function getFileCacheKey(file: File | ArrayBuffer): string | null {
  if (file instanceof File) {
    return `${file.name}:${file.size}:${file.lastModified}`;
  }
  return null;
}

/**
 * Extracts text items across all pages in parallel and reconstructs visual lines
 * sorted top-to-bottom, left-to-right.
 */
export async function extractPdfLines(file: File | ArrayBuffer): Promise<{
  lines: PdfLine[];
  rawText: string;
  pageCount: number;
}> {
  const cacheKey = getFileCacheKey(file);
  if (cacheKey) {
    const cached = pdfLinesCache.get(cacheKey);
    if (cached) return cached;
  }

  const buffer = file instanceof File ? await file.arrayBuffer() : file;
  const loadingTask = pdfjsLib.getDocument({
    data: new Uint8Array(buffer),
    worker: getSharedPdfWorker(),
  });
  const pdf = await loadingTask.promise;

  const pagePromises = Array.from({ length: pdf.numPages }, async (_, idx) => {
    const p = idx + 1;
    const page = await pdf.getPage(p);
    const textContent = await page.getTextContent();
    const items: PdfTextItem[] = [];

    for (const rawItem of textContent.items) {
      if ('str' in rawItem && rawItem.str.trim().length > 0) {
        const tx = rawItem.transform;
        items.push({
          text: rawItem.str,
          x: tx[4],
          y: tx[5],
          width: rawItem.width ?? 0,
          height: rawItem.height ?? 0,
          page: p,
        });
      }
    }

    items.sort((a, b) => {
      if (Math.abs(b.y - a.y) > 3.5) {
        return b.y - a.y;
      }
      return a.x - b.x;
    });

    const pageLines: PdfLine[] = [];
    for (const item of items) {
      const existing = pageLines.find((line) => Math.abs(line.y - item.y) <= 4);
      if (existing) {
        existing.items.push(item);
      } else {
        pageLines.push({
          y: item.y,
          page: p,
          items: [item],
          text: '',
        });
      }
    }

    for (const line of pageLines) {
      line.items.sort((a, b) => a.x - b.x);
      line.text = line.items
        .map((i) => i.text.trim())
        .filter(Boolean)
        .join(' ');
    }

    return pageLines;
  });

  const perPageLines = await Promise.all(pagePromises);
  const allLines = perPageLines.flat();
  const rawText = allLines.map((l) => l.text).join('\n');

  const result = { lines: allLines, rawText, pageCount: pdf.numPages };
  if (cacheKey) {
    pdfLinesCache.set(cacheKey, result);
  }
  return result;
}

const pdfImagesCache = new Map<
  string,
  Array<{ mimeType: string; data: string; width: number; height: number }>
>();

/**
 * Renders PDF pages in parallel to optimally sized JPEG images
 * for ultra-fast network transfer and Gemini multimodal extraction.
 */
export async function renderPdfPagesToImages(
  file: File,
  maxWidth = 1400,
  maxPages = 10
): Promise<Array<{ mimeType: string; data: string; width: number; height: number }>> {
  const fileKey = getFileCacheKey(file);
  const imgCacheKey = fileKey ? `${fileKey}:${maxWidth}:${maxPages}` : null;
  if (imgCacheKey) {
    const cached = pdfImagesCache.get(imgCacheKey);
    if (cached) return cached;
  }

  const buffer = await file.arrayBuffer();
  const loadingTask = pdfjsLib.getDocument({
    data: new Uint8Array(buffer),
    worker: getSharedPdfWorker(),
  });
  const pdf = await loadingTask.promise;

  const pagesToRender = Math.min(pdf.numPages, maxPages);

  const renderPromises = Array.from({ length: pagesToRender }, async (_, idx) => {
    const p = idx + 1;
    const page = await pdf.getPage(p);
    const baseViewport = page.getViewport({ scale: 1 });
    const targetScale = Math.min(1.65, Math.max(1.1, maxWidth / baseViewport.width));
    const viewport = page.getViewport({ scale: targetScale });

    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d');
    if (!context) return null;

    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);

    await page.render({
      canvasContext: context,
      viewport,
      canvas,
    }).promise;

    const dataUrl = canvas.toDataURL('image/jpeg', 0.62);
    const base64 = dataUrl.split(',')[1] ?? '';
    return {
      mimeType: 'image/jpeg',
      data: base64,
      width: canvas.width,
      height: canvas.height,
    };
  });

  const results = await Promise.all(renderPromises);
  const filtered = results.filter(
    (r): r is { mimeType: string; data: string; width: number; height: number } => r !== null
  );
  if (imgCacheKey) {
    pdfImagesCache.set(imgCacheKey, filtered);
  }
  return filtered;
}
