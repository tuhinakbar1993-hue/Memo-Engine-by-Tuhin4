import { renderAsync } from 'docx-preview';
import html2canvas from 'html2canvas-pro';
import { jsPDF } from 'jspdf';
import type {
  MemoDocType,
  MemoTemplateRecord,
  ResolvedICMemo,
  ResolvedStageMemo,
} from '../types/memo';
import { populateDocxTemplate } from './docxTemplateEngine';
import { generateICMemoDocxBlob } from './icMemoDocx';
import { generateStageMemoDocxBlob } from './stageMemoDocx';

const PX_TO_PT = 72 / 96;
const A4_RATIO = 841.92 / 595.32;

async function waitForImagesAndFonts(root: HTMLElement): Promise<void> {
  if ('fonts' in document && document.fonts?.ready) {
    await document.fonts.ready.catch(() => undefined);
  }
  const imgs = Array.from(root.querySelectorAll('img'));
  await Promise.all(
    imgs.map(
      (img) =>
        new Promise<void>((resolve) => {
          if (img.complete && img.naturalWidth > 0) {
            resolve();
            return;
          }
          const finish = () => resolve();
          img.addEventListener('load', finish, { once: true });
          img.addEventListener('error', finish, { once: true });
          window.setTimeout(finish, 2500);
        })
    )
  );
  // Allow one animation frame for layout settlement after fonts/images settle
  await new Promise<void>((resolve) => window.requestAnimationFrame(() => resolve()));
}

/**
 * Renders a .docx Blob into the supplied DOM container using docx-preview,
 * preserving tables, nested tables, fonts, margins, orientation, borders,
 * shading, headers, footers, page numbers, and embedded images.
 */
export async function renderDocxBlobToContainer(
  docxBlob: Blob,
  bodyContainer: HTMLElement,
  styleContainer?: HTMLElement,
  options?: { forPdfCapture?: boolean }
): Promise<void> {
  bodyContainer.innerHTML = '';
  if (styleContainer) {
    styleContainer.innerHTML = '';
  }

  await renderAsync(docxBlob, bodyContainer, styleContainer, {
    className: 'docx',
    inWrapper: true,
    hideWrapperOnPrint: false,
    ignoreWidth: false,
    ignoreHeight: false,
    ignoreFonts: false,
    breakPages: true,
    ignoreLastRenderedPageBreak: false,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
    renderEndnotes: true,
    useBase64URL: true,
    experimental: false,
    renderChanges: false,
    renderComments: false,
    renderAltChunks: true,
  });

  if (options?.forPdfCapture) {
    const wrapper = bodyContainer.querySelector<HTMLElement>('.docx-wrapper');
    if (wrapper) {
      wrapper.style.background = '#ffffff';
      wrapper.style.padding = '0';
      wrapper.style.display = 'block';
    }
    const sections = Array.from(bodyContainer.querySelectorAll<HTMLElement>('section.docx'));
    for (const sec of sections) {
      sec.style.boxShadow = 'none';
      sec.style.margin = '0';
      sec.style.background = '#ffffff';
    }
  }

  await waitForImagesAndFonts(bodyContainer);
}

/**
 * Generates an editable .docx Blob using strict per-memo-type template isolation:
 * - If an uploaded template exists for `docType`, populates that exact uploaded template.
 * - Otherwise uses the dedicated built-in engine for IC Memo or the stage engine for MEC/Board/Sanction.
 */
export async function generateSynchronizedDocxBlob(
  docType: MemoDocType,
  memo: ResolvedICMemo,
  activeTemplate: MemoTemplateRecord | null,
  stageMemo?: ResolvedStageMemo | null
): Promise<Blob> {
  if (activeTemplate) {
    return populateDocxTemplate(activeTemplate.bytes, memo, docType, stageMemo);
  }

  if (docType === 'ic') {
    return generateICMemoDocxBlob(memo);
  }

  if (!stageMemo) {
    throw new Error(`Stage data missing for ${docType.toUpperCase()} export.`);
  }

  return generateStageMemoDocxBlob(stageMemo);
}

/**
 * Renders a pre-generated (and verified) .docx Blob directly into a multi-page PDF Blob,
 * preserving original Word page sizes, orientation, margins, headers/footers,
 * tables, fonts, colors, and graphical elements.
 */
export async function generatePdfFromDocxBlob(docxBlob: Blob): Promise<Blob> {
  const host = document.createElement('div');
  host.style.position = 'fixed';
  host.style.left = '-14000px';
  host.style.top = '0';
  host.style.background = '#ffffff';
  host.style.color = '#000000';
  host.style.zIndex = '-1';
  host.style.pointerEvents = 'none';

  const styleContainer = document.createElement('div');
  const bodyContainer = document.createElement('div');
  host.appendChild(styleContainer);
  host.appendChild(bodyContainer);
  document.body.appendChild(host);

  try {
    await renderDocxBlobToContainer(docxBlob, bodyContainer, styleContainer, {
      forPdfCapture: true,
    });

    const sectionEls = Array.from(bodyContainer.querySelectorAll<HTMLElement>('section.docx'));
    const targets = sectionEls.length > 0 ? sectionEls : [bodyContainer];

    let pdf: jsPDF | null = null;

    for (const sectionEl of targets) {
      const rect = sectionEl.getBoundingClientRect();
      const computed = window.getComputedStyle(sectionEl);
      const secWidthPx = Math.max(200, Math.round(rect.width || sectionEl.offsetWidth || 794));
      const secHeightPx = Math.max(200, Math.round(rect.height || sectionEl.scrollHeight || 1123));

      const minHeightPxRaw = parseFloat(computed.minHeight || '0');
      const targetPageHeightPx =
        Number.isFinite(minHeightPxRaw) && minHeightPxRaw > 200
          ? minHeightPxRaw
          : Math.round(secWidthPx * A4_RATIO);

      const canvas = await html2canvas(sectionEl, {
        scale: 2,
        backgroundColor: '#ffffff',
        useCORS: true,
        logging: false,
        width: secWidthPx,
        height: secHeightPx,
        windowWidth: secWidthPx,
        windowHeight: secHeightPx,
      });

      const scaleFactor = canvas.width / secWidthPx;
      const singlePageWidthPt = secWidthPx * PX_TO_PT;

      // Single-page section or nearly exact page height
      if (secHeightPx <= targetPageHeightPx * 1.06) {
        const pageHeightPt = Math.max(targetPageHeightPx, secHeightPx) * PX_TO_PT;
        const orientation = singlePageWidthPt > pageHeightPt ? 'landscape' : 'portrait';

        if (!pdf) {
          pdf = new jsPDF({
            orientation,
            unit: 'pt',
            format: [singlePageWidthPt, pageHeightPt],
            compress: true,
          });
        } else {
          pdf.addPage([singlePageWidthPt, pageHeightPt], orientation);
        }

        const imgData = canvas.toDataURL('image/jpeg', 0.96);
        pdf.addImage(imgData, 'JPEG', 0, 0, singlePageWidthPt, secHeightPx * PX_TO_PT, undefined, 'FAST');
        continue;
      }

      // Tall unsplit Word section: slice crisply into page-sized strips
      const sliceCanvasHeight = Math.max(100, Math.round(targetPageHeightPx * scaleFactor));
      const pageHeightPt = targetPageHeightPx * PX_TO_PT;
      const orientation = singlePageWidthPt > pageHeightPt ? 'landscape' : 'portrait';

      let yOffsetCanvas = 0;
      while (yOffsetCanvas < canvas.height) {
        const remainingCanvasHeight = Math.min(sliceCanvasHeight, canvas.height - yOffsetCanvas);
        const sliceCanvas = document.createElement('canvas');
        sliceCanvas.width = canvas.width;
        sliceCanvas.height = sliceCanvasHeight;

        const ctx = sliceCanvas.getContext('2d');
        if (ctx) {
          ctx.fillStyle = '#ffffff';
          ctx.fillRect(0, 0, sliceCanvas.width, sliceCanvas.height);
          ctx.drawImage(
            canvas,
            0,
            yOffsetCanvas,
            canvas.width,
            remainingCanvasHeight,
            0,
            0,
            canvas.width,
            remainingCanvasHeight
          );
        }

        if (!pdf) {
          pdf = new jsPDF({
            orientation,
            unit: 'pt',
            format: [singlePageWidthPt, pageHeightPt],
            compress: true,
          });
        } else {
          pdf.addPage([singlePageWidthPt, pageHeightPt], orientation);
        }

        const sliceData = sliceCanvas.toDataURL('image/jpeg', 0.96);
        pdf.addImage(sliceData, 'JPEG', 0, 0, singlePageWidthPt, pageHeightPt, undefined, 'FAST');
        yOffsetCanvas += sliceCanvasHeight;
      }
    }

    if (!pdf) {
      throw new Error('Failed to generate PDF pages from document.');
    }

    return pdf.output('blob');
  } finally {
    host.remove();
  }
}

/**
 * Converts a synchronized .docx output directly into a multi-page PDF Blob,
 * preserving original Word page sizes, orientation, margins, headers/footers,
 * tables, fonts, colors, and graphical elements.
 */
export async function generateMemoPdfBlob(
  docType: MemoDocType,
  memo: ResolvedICMemo,
  activeTemplate: MemoTemplateRecord | null,
  stageMemo?: ResolvedStageMemo | null
): Promise<Blob> {
  const docxBlob = await generateSynchronizedDocxBlob(docType, memo, activeTemplate, stageMemo);
  return generatePdfFromDocxBlob(docxBlob);
}
