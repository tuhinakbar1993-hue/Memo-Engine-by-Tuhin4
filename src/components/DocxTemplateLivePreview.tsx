import React, { useEffect, useRef, useState } from 'react';
import { Loader2 } from 'lucide-react';
import {
  populateDocxTemplate,
  renderTemplatePreviewBlocks,
} from '../exporters/docxTemplateEngine';
import { renderDocxBlobToContainer } from '../exporters/pdfExporter';
import type {
  MemoDocType,
  MemoTemplateRecord,
  ResolvedICMemo,
  ResolvedStageMemo,
} from '../types/memo';

interface DocxTemplateLivePreviewProps {
  template: MemoTemplateRecord;
  memo: ResolvedICMemo;
  docType: MemoDocType;
  stageMemo?: ResolvedStageMemo | null;
}

export const DocxTemplateLivePreview: React.FC<DocxTemplateLivePreviewProps> = ({
  template,
  memo,
  docType,
  stageMemo,
}) => {
  const styleContainerRef = useRef<HTMLDivElement | null>(null);
  const bodyContainerRef = useRef<HTMLDivElement | null>(null);
  const hasRenderedOnceRef = useRef(false);
  const renderSeqRef = useRef(0);

  const [rendering, setRendering] = useState(true);
  const [fallbackError, setFallbackError] = useState<string | null>(null);

  useEffect(() => {
    hasRenderedOnceRef.current = false;
  }, [template]);

  useEffect(() => {
    let cancelled = false;
    const seq = ++renderSeqRef.current;
    const delayMs = hasRenderedOnceRef.current ? 180 : 0;

    const timer = window.setTimeout(async () => {
      const bodyEl = bodyContainerRef.current;
      const styleEl = styleContainerRef.current;
      if (!bodyEl || !styleEl || cancelled) return;

      setRendering(true);
      try {
        const populatedBlob = await populateDocxTemplate(
          template.bytes,
          memo,
          docType,
          stageMemo
        );
        if (cancelled || seq !== renderSeqRef.current) return;

        await renderDocxBlobToContainer(populatedBlob, bodyEl, styleEl);
        if (cancelled || seq !== renderSeqRef.current) return;

        hasRenderedOnceRef.current = true;
        setFallbackError(null);
      } catch (err) {
        if (cancelled || seq !== renderSeqRef.current) return;
        setFallbackError(
          err instanceof Error
            ? err.message
            : 'Unable to render exact Word preview.'
        );
      } finally {
        if (!cancelled && seq === renderSeqRef.current) {
          setRendering(false);
        }
      }
    }, delayMs);

    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, [template, memo, docType, stageMemo]);

  return (
    <div className="bg-slate-200/70 p-4 sm:p-8 rounded-lg space-y-3">
      <div className="max-w-[920px] mx-auto bg-emerald-50 border border-emerald-200 rounded-md px-4 py-2 text-xs text-emerald-950 flex flex-wrap items-center justify-between gap-2">
        <span>
          <span className="font-semibold">Active Uploaded Template:</span>{' '}
          {template.fileName} — exact Word page layout, margins, fonts, table
          dimensions, headers &amp; footers preserved.
        </span>
        {rendering && (
          <span className="inline-flex items-center gap-1.5 text-[11px] font-medium text-emerald-900 shrink-0">
            <Loader2 className="w-3.5 h-3.5 animate-spin" />
            <span>Syncing preview...</span>
          </span>
        )}
      </div>

      {fallbackError ? (
        <div className="ic-memo-sheet max-w-[794px] mx-auto bg-white shadow-sm border border-slate-300 px-12 py-8 space-y-2.5 text-black text-[10pt]">
          <p className="text-xs text-amber-800 bg-amber-50 border border-amber-200 rounded px-3 py-1.5 font-sans">
            Showing simplified preview ({fallbackError}). Downloaded .docx still preserves original template XML.
          </p>
          {renderTemplatePreviewBlocks(template, memo, stageMemo).map((block, idx) => {
            if (block.type === 'paragraph') {
              return (
                <p key={idx} className="whitespace-pre-line leading-snug">
                  {block.text}
                </p>
              );
            }
            return (
              <table
                key={idx}
                className="w-full table-fixed border-collapse border border-black text-[9.5pt]"
              >
                <tbody>
                  {(block.rows || []).map((row, rIdx) => (
                    <tr key={rIdx}>
                      {row.map((cell, cIdx) => (
                        <td
                          key={cIdx}
                          className="border border-black px-2 py-1 align-middle whitespace-pre-line"
                        >
                          {cell}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            );
          })}
        </div>
      ) : (
        <div className="overflow-x-auto">
          <div ref={styleContainerRef} />
          <div
            ref={bodyContainerRef}
            className="docx-live-preview-host min-h-[240px]"
          />
        </div>
      )}
    </div>
  );
};
