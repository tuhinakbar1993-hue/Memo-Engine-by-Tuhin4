import {
  AlignmentType,
  Document,
  Packer,
  Paragraph,
  Table,
  TableLayoutType,
  TableRow,
  WidthType,
} from 'docx';
import type { ResolvedStageMemo } from '../types/memo';
import {
  A4_HEIGHT_TWIPS,
  A4_WIDTH_TWIPS,
  ALL_THIN_BORDERS,
  BODY_WIDTH_TWIPS,
  MARGIN_LR_TWIPS,
  SZ_BODY,
  SZ_META,
  buildBankHeaderAndFooter,
  buildICMemoCoreBlocks,
  cell,
  p,
  pctWidths,
  simpleCell,
  txt,
  valOrToken,
} from './icMemoDocx';

function renderNarrativeParagraphs(
  heading: string,
  text: string,
  opts: { before?: number; after?: number; fallback?: string } = {}
): Paragraph[] {
  const blocks: Paragraph[] = [
    p([txt(heading, { bold: true, underline: true })], {
      before: opts.before ?? 45,
      after: 15,
    }),
  ];
  const content = (text || '').trim() || opts.fallback || '';
  if (content) {
    for (const line of content.split('\n').filter((l) => l.trim())) {
      blocks.push(
        p([txt(line.trim())], {
          indentLeft: 180,
          before: 12,
          after: opts.after ?? 16,
          line: 245,
        })
      );
    }
  }
  return blocks;
}

function buildStageMetadataTable(
  rows: Array<{ label: string; value: string; boldValue?: boolean }>
): Array<Paragraph | Table> {
  const metaWidth = 5600;
  const metaCols = pctWidths([34, 5, 61], metaWidth);
  return [
    new Table({
      width: { size: metaWidth, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: metaCols,
      rows: rows.map(
        (r) =>
          new TableRow({
            cantSplit: true,
            children: [
              simpleCell(r.label, {
                bold: true,
                size: SZ_META,
                widthTwips: metaCols[0],
                borders: ALL_THIN_BORDERS,
                marginTB: 22,
              }),
              simpleCell(':', {
                bold: true,
                size: SZ_META,
                align: AlignmentType.CENTER,
                widthTwips: metaCols[1],
                borders: ALL_THIN_BORDERS,
                marginTB: 22,
              }),
              simpleCell(r.value, {
                bold: r.boldValue ?? true,
                size: SZ_BODY,
                widthTwips: metaCols[2],
                borders: ALL_THIN_BORDERS,
                marginTB: 22,
              }),
            ],
          })
      ),
    }),
    p([], { before: 20, after: 20, line: 140 }),
  ];
}

function buildSignatoriesAndPlacement(placementText: string, subtitleText: string): Array<Paragraph | Table> {
  const sigCols = pctWidths([50, 50]);
  return [
    p([], { before: 50, after: 30 }),
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: [BODY_WIDTH_TWIPS],
      rows: [
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell(placementText, {
              bold: true,
              align: AlignmentType.CENTER,
              size: SZ_BODY,
              widthTwips: BODY_WIDTH_TWIPS,
              borders: ALL_THIN_BORDERS,
              marginTB: 35,
            }),
          ],
        }),
      ],
    }),
    p([txt('Union Bank PLC Unit', { bold: true })], { before: 60, after: 10 }),
    p([txt('Special Asset Management Division', { bold: true })], { before: 0, after: 45 }),
    p([txt(subtitleText, { bold: true })], { align: AlignmentType.CENTER, before: 35, after: 95 }),
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: sigCols,
      rows: [
        new TableRow({
          cantSplit: true,
          children: [
            cell(
              [
                p([txt('(Shahidur Rahman khan)', { bold: true })], { before: 180, after: 10 }),
                p([txt('Deputy Managing Director & CRO')], { before: 0, after: 10 }),
                p([txt('Sammilito Islami Bank PLC.')], { before: 0, after: 10 }),
              ],
              { widthTwips: sigCols[0], borders: ALL_THIN_BORDERS, marginLR: 100 }
            ),
            cell(
              [
                p([txt('(Ahmed Zafrul Hasan)', { bold: true })], { before: 180, after: 10 }),
                p([txt('Additional Managing Director & COO')], { before: 0, after: 10 }),
                p([txt('Sammilito Islami Bank PLC.')], { before: 0, after: 10 }),
              ],
              { widthTwips: sigCols[1], borders: ALL_THIN_BORDERS, marginLR: 100 }
            ),
          ],
        }),
      ],
    }),
    p([txt('(Abedur Rahman Sikder)', { bold: true })], { align: AlignmentType.CENTER, before: 220, after: 10 }),
    p([txt('Managing Director')], { align: AlignmentType.CENTER, before: 0, after: 10 }),
    p([txt('Sammilito Islami Bank PLC.')], { align: AlignmentType.CENTER, before: 0, after: 10 }),
    p([txt('(Transferee of the assets and liabilities of')], { align: AlignmentType.CENTER, before: 0, after: 10 }),
    p([txt('EXIM Bank, SIBL, FSIB, Global Islami')], { align: AlignmentType.CENTER, before: 0, after: 10 }),
    p([txt('Bank and Union Bank)')], { align: AlignmentType.CENTER, before: 0, after: 10 }),
  ];
}

export async function generateStageMemoDocxBlob(stageMemo: ResolvedStageMemo): Promise<Blob> {
  const { docType, icMemo, stageInputs, stageSections } = stageMemo;
  const ph = icMemo.placeholders;
  const v = (code: string) => valOrToken(ph, code);

  const badgeLabel =
    docType === 'mec'
      ? 'MEC Memo'
      : docType === 'board'
        ? 'Board Memo'
        : 'Sanction Letter';

  const { header, footer } = await buildBankHeaderAndFooter(badgeLabel);
  const body: Array<Paragraph | Table> = [];

  if (docType === 'mec') {
    body.push(
      ...buildStageMetadataTable([
        { label: 'Dated', value: stageInputs.mecMemoDate || stageInputs.icMemoDate || v('#1#') },
        { label: 'Chronologic Order of the MEC Memo', value: stageInputs.mecMemoNumber || '', boldValue: Boolean(stageInputs.mecMemoNumber) },
        { label: 'Chronologic Order of the IC Memo', value: stageInputs.icMemoNumber || '', boldValue: Boolean(stageInputs.icMemoNumber) },
        { label: 'IC Memo Date', value: stageInputs.icMemoDate || '', boldValue: Boolean(stageInputs.icMemoDate) },
        { label: 'To', value: 'Management Executive Committee (MEC)' },
        { label: 'From', value: 'SAMD' },
      ])
    );

    // Full retained IC Memo core blocks + Section 16 Recommendation of the Investment Committee (IC)
    body.push(
      ...buildICMemoCoreBlocks(icMemo, {
        includeClosingBlock: false,
        skipMetadataTable: true,
        insertedAfterSec15: [
          {
            sectionNumber: 16,
            heading: 'Recommendation of the Investment Committee (IC)',
            text: stageSections.mecIcRecommendation,
          },
        ],
        particularsSectionNumber: 17,
        termsSectionNumber: 18,
      })
    );

    body.push(
      ...buildSignatoriesAndPlacement(
        'The proposal is hereby placed before the Management Executive Committee (MEC) for kind recommendation/approval.',
        'Submitted before the Management Executive Committee (MEC) for kind decision.'
      )
    );
  } else if (docType === 'board') {
    body.push(
      ...buildStageMetadataTable([
        { label: 'Dated', value: stageInputs.boardMemoDate || stageInputs.mecMemoDate || v('#1#') },
        { label: 'Chronologic Order of the Board Memo', value: stageInputs.boardMemoNumber || '', boldValue: Boolean(stageInputs.boardMemoNumber) },
        { label: 'Chronologic Order of the MEC Memo', value: stageInputs.mecMemoNumber || '', boldValue: Boolean(stageInputs.mecMemoNumber) },
        { label: 'MEC Memo Date', value: stageInputs.mecMemoDate || '', boldValue: Boolean(stageInputs.mecMemoDate) },
        { label: 'Chronologic Order of the IC Memo & Date', value: `${stageInputs.icMemoNumber || '-'} dated ${stageInputs.icMemoDate || '-'}` },
        { label: 'To', value: 'Board of Directors' },
        { label: 'From', value: 'Management Executive Committee (MEC) / SAMD' },
      ])
    );

    body.push(
      ...buildICMemoCoreBlocks(icMemo, {
        includeClosingBlock: false,
        skipMetadataTable: true,
        insertedAfterSec15: [
          {
            sectionNumber: 16,
            heading: 'Recommendation of the Investment Committee (IC)',
            text: stageSections.mecIcRecommendation,
          },
          {
            sectionNumber: 17,
            heading: 'Management Executive Committee (MEC) Recommendation',
            text: stageSections.boardMecRecommendation,
          },
        ],
        particularsSectionNumber: 18,
        termsSectionNumber: 19,
      })
    );

    body.push(
      ...buildSignatoriesAndPlacement(
        'Submitted before the Honourable Board of Directors for kind consideration and approval.',
        'Submitted for kind approval of the Board of Directors.'
      )
    );
  } else {
    // Sanction Letter
    body.push(
      ...buildStageMetadataTable([
        { label: 'Sanction Ref. No.', value: stageInputs.sanctionRefNumber || '', boldValue: Boolean(stageInputs.sanctionRefNumber) },
        { label: 'Sanction Date', value: stageInputs.sanctionDate || stageInputs.boardMemoDate || v('#1#') },
        { label: 'Chronologic Order of the Board Memo', value: stageInputs.boardMemoNumber || '', boldValue: Boolean(stageInputs.boardMemoNumber) },
        { label: 'Board Memo Date', value: stageInputs.boardMemoDate || '', boldValue: Boolean(stageInputs.boardMemoDate) },
        { label: 'Chronologic Order of the MEC Memo & Date', value: `${stageInputs.mecMemoNumber || '-'} dated ${stageInputs.mecMemoDate || '-'}` },
        { label: 'Chronologic Order of the IC Memo & Date', value: `${stageInputs.icMemoNumber || '-'} dated ${stageInputs.icMemoDate || '-'}` },
        { label: 'To', value: `The Head of Branch, ${v('#3#')} Branch` },
        { label: 'From', value: 'Special Asset Management Division (SAMD), Head Office' },
      ])
    );

    body.push(
      ...buildICMemoCoreBlocks(icMemo, {
        includeClosingBlock: false,
        skipMetadataTable: true,
      })
    );

    if (stageSections.mecIcRecommendation.trim()) {
      body.push(
        ...renderNarrativeParagraphs(
          'Recommendation of the Investment Committee (IC) (Carried Forward):',
          stageSections.mecIcRecommendation,
          { before: 50 }
        )
      );
    }
    if (stageSections.boardMecRecommendation.trim()) {
      body.push(
        ...renderNarrativeParagraphs(
          'Recommendation of the Management Executive Committee (MEC) (Carried Forward):',
          stageSections.boardMecRecommendation,
          { before: 40 }
        )
      );
    }

    body.push(
      ...renderNarrativeParagraphs(
        'AI Cross-Verified Facility & Liability Confirmation:',
        stageSections.sanctionCrossVerifiedSummary,
        { before: 55 }
      ),
      ...buildSignatoriesAndPlacement(
        'Issued under Head Office Sanction Authority. Strict compliance with all stipulated terms and BRPD circulars is mandatory.',
        'Conveyed for immediate execution and compliance by the Branch.'
      )
    );
  }

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            size: {
              width: A4_WIDTH_TWIPS,
              height: A4_HEIGHT_TWIPS,
            },
            margin: {
              top: 1520,
              bottom: 680,
              left: MARGIN_LR_TWIPS,
              right: MARGIN_LR_TWIPS,
              header: 300,
              footer: 280,
            },
          },
        },
        headers: { default: header },
        footers: { default: footer },
        children: body,
      },
    ],
  });

  return await Packer.toBlob(doc);
}
