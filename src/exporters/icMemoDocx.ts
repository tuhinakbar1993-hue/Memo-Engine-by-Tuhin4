import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  ImageRun,
  type ITableCellBorders,
  Packer,
  PageNumber,
  Paragraph,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  type TableVerticalAlign,
  TextRun,
  UnderlineType,
  VerticalAlign,
  WidthType,
} from 'docx';
import type { ResolvedICMemo } from '../types/memo';

export const FONT = 'Times New Roman';
export const SZ_BODY = 20; // 10pt
export const SZ_PARTICULARS = 19; // 9.5pt
export const SZ_META = 21; // 10.5pt
export const SZ_SMALL = 18; // 9pt
export const SZ_TITLE_BANK = 28; // 14pt
export const SZ_TITLE_SUB = 26; // 13pt
export const SZ_TITLE_UNIT = 22; // 11pt
export const SZ_TITLE_WAIVER = 24; // 12pt

// A4 dimensions in twips (595.32pt x 841.92pt) — strictly preserve existing page margins
export const A4_WIDTH_TWIPS = 11906;
export const A4_HEIGHT_TWIPS = 16838;
export const MARGIN_LR_TWIPS = 1440; // 1 inch (72pt) matching reference IC_Memo.pdf
export const BODY_WIDTH_TWIPS = A4_WIDTH_TWIPS - MARGIN_LR_TWIPS * 2; // 9026 twips

export const THIN_BORDER = {
  style: BorderStyle.SINGLE,
  size: 4,
  color: '000000',
};

export const DOTTED_BORDER = {
  style: BorderStyle.DOTTED,
  size: 4,
  color: '444444',
};

export const NO_BORDER = {
  style: BorderStyle.NONE,
  size: 0,
  color: 'FFFFFF',
};

export const ALL_THIN_BORDERS: ITableCellBorders = {
  top: THIN_BORDER,
  bottom: THIN_BORDER,
  left: THIN_BORDER,
  right: THIN_BORDER,
};

export const ALL_DOTTED_BORDERS: ITableCellBorders = {
  top: DOTTED_BORDER,
  bottom: DOTTED_BORDER,
  left: DOTTED_BORDER,
  right: DOTTED_BORDER,
};

export const ALL_NO_BORDERS: ITableCellBorders = {
  top: NO_BORDER,
  bottom: NO_BORDER,
  left: NO_BORDER,
  right: NO_BORDER,
};

export function txt(
  text: string,
  opts: {
    bold?: boolean;
    underline?: boolean;
    size?: number;
    italics?: boolean;
    superScript?: boolean;
  } = {}
): TextRun {
  return new TextRun({
    text,
    font: FONT,
    size: opts.size ?? SZ_BODY,
    bold: opts.bold,
    italics: opts.italics,
    superScript: opts.superScript,
    underline: opts.underline ? { type: UnderlineType.SINGLE } : undefined,
  });
}

export function p(
  children: TextRun[],
  opts: {
    align?: (typeof AlignmentType)[keyof typeof AlignmentType];
    before?: number;
    after?: number;
    indentLeft?: number;
    pageBreakBefore?: boolean;
    line?: number;
  } = {}
): Paragraph {
  return new Paragraph({
    alignment: opts.align ?? AlignmentType.LEFT,
    spacing: {
      before: opts.before ?? 24,
      after: opts.after ?? 24,
      line: opts.line ?? 240,
    },
    indent: opts.indentLeft ? { left: opts.indentLeft } : undefined,
    pageBreakBefore: opts.pageBreakBefore,
    children,
  });
}

export function cell(
  paragraphs: Array<Paragraph | Table>,
  opts: {
    widthPct?: number;
    widthTwips?: number;
    rowSpan?: number;
    columnSpan?: number;
    borders?: ITableCellBorders;
    vAlign?: TableVerticalAlign;
    shading?: string;
    marginTB?: number;
    marginLR?: number;
  } = {}
): TableCell {
  const width = opts.widthTwips
    ? { size: opts.widthTwips, type: WidthType.DXA }
    : opts.widthPct
      ? { size: opts.widthPct, type: WidthType.PERCENTAGE }
      : undefined;

  return new TableCell({
    width,
    rowSpan: opts.rowSpan,
    columnSpan: opts.columnSpan,
    borders: opts.borders ?? ALL_DOTTED_BORDERS,
    verticalAlign: opts.vAlign ?? VerticalAlign.CENTER,
    shading: opts.shading ? { fill: opts.shading } : undefined,
    margins: {
      top: opts.marginTB ?? 36,
      bottom: opts.marginTB ?? 36,
      left: opts.marginLR ?? 76,
      right: opts.marginLR ?? 76,
    },
    children: paragraphs,
  });
}

export function simpleCell(
  text: string,
  opts: {
    bold?: boolean;
    align?: (typeof AlignmentType)[keyof typeof AlignmentType];
    size?: number;
    widthPct?: number;
    widthTwips?: number;
    rowSpan?: number;
    columnSpan?: number;
    borders?: ITableCellBorders;
    vAlign?: TableVerticalAlign;
    shading?: string;
    marginTB?: number;
    marginLR?: number;
  } = {}
): TableCell {
  const lines = (text || '').split('\n');
  const paras = lines.map((line) =>
    p([txt(line, { bold: opts.bold, size: opts.size ?? SZ_BODY })], {
      align: opts.align ?? AlignmentType.LEFT,
      before: 10,
      after: 10,
      line: 225,
    })
  );
  return cell(paras, {
    widthPct: opts.widthPct,
    widthTwips: opts.widthTwips,
    rowSpan: opts.rowSpan,
    columnSpan: opts.columnSpan,
    borders: opts.borders,
    vAlign: opts.vAlign,
    shading: opts.shading,
    marginTB: opts.marginTB,
    marginLR: opts.marginLR,
  });
}

export function pctWidths(pcts: number[], totalTwips = BODY_WIDTH_TWIPS): number[] {
  const widths = pcts.map((pct) => Math.floor((pct / 100) * totalTwips));
  const diff = totalTwips - widths.reduce((a, b) => a + b, 0);
  if (widths.length > 0) widths[widths.length - 1] += diff;
  return widths;
}

export async function fetchImageBytes(url: string): Promise<Uint8Array | null> {
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    const contentType = res.headers.get('content-type') ?? '';
    if (/text\/html/i.test(contentType)) return null;
    const buf = await res.arrayBuffer();
    return new Uint8Array(buf);
  } catch {
    return null;
  }
}

export function valOrToken(ph: Record<string, string>, code: string): string {
  const v = ph[code];
  return v && v.trim() !== '' ? v.trim() : code;
}

export async function buildBankHeaderAndFooter(
  badgeLabel = 'Exit with Waiver'
): Promise<{ header: Header; footer: Footer }> {
  const [logoBytes, bismillahBytes] = await Promise.all([
    fetchImageBytes('/assets/sib_logo.jpg'),
    fetchImageBytes('/assets/bismillah.jpg'),
  ]);

  // Header rendered identically across all A4 pages (width extends slightly into left/right margins just like IC Memo.pdf)
  const HEADER_WIDTH_TWIPS = 10466;
  const headerCols = pctWidths([16, 65, 19], HEADER_WIDTH_TWIPS);
  const badgeBoxWidth = headerCols[2] - 120;
  const headerChildren: Array<Paragraph | Table> = [];
  if (bismillahBytes) {
    headerChildren.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 20 },
        children: [
          new ImageRun({
            data: bismillahBytes,
            type: 'jpg',
            transformation: { width: 175, height: 20 },
          }),
        ],
      })
    );
  }

  headerChildren.push(
    new Table({
      width: { size: HEADER_WIDTH_TWIPS, type: WidthType.DXA },
      alignment: AlignmentType.CENTER,
      layout: TableLayoutType.FIXED,
      columnWidths: headerCols,
      rows: [
        new TableRow({
          cantSplit: true,
          children: [
            cell(
              logoBytes
                ? [
                    new Paragraph({
                      alignment: AlignmentType.LEFT,
                      spacing: { before: 0, after: 0 },
                      children: [
                        new ImageRun({
                          data: logoBytes,
                          type: 'jpg',
                          transformation: { width: 54, height: 61 },
                        }),
                      ],
                    }),
                  ]
                : [p([txt('SIB PLC', { bold: true })])],
              { widthTwips: headerCols[0], borders: ALL_NO_BORDERS, marginTB: 0, vAlign: VerticalAlign.TOP }
            ),
            cell(
              [
                p([txt('SAMMILITO ISLAMI BANK PLC.', { underline: true, size: SZ_TITLE_BANK })], {
                  align: AlignmentType.CENTER,
                  before: 0,
                  after: 10,
                  line: 240,
                }),
                p(
                  [
                    txt('(', { underline: true, size: SZ_SMALL }),
                    txt('State-owned Shariah-based Bank', { underline: true, size: SZ_TITLE_SUB }),
                    txt(')', { underline: true, size: SZ_SMALL }),
                  ],
                  {
                    align: AlignmentType.CENTER,
                    before: 0,
                    after: 10,
                    line: 240,
                  }
                ),
                p([txt('UNION BANK UNIT', { underline: true, size: SZ_TITLE_UNIT })], {
                  align: AlignmentType.CENTER,
                  before: 0,
                  after: 10,
                  line: 240,
                }),
                p(
                  [
                    txt('Special Asset Management Division Head Office, Dhaka', {
                      underline: true,
                      size: SZ_SMALL,
                    }),
                  ],
                  {
                    align: AlignmentType.CENTER,
                    before: 0,
                    after: 10,
                    line: 220,
                  }
                ),
              ],
              { widthTwips: headerCols[1], borders: ALL_NO_BORDERS, marginTB: 0, vAlign: VerticalAlign.TOP }
            ),
            cell(
              [
                new Table({
                  width: { size: badgeBoxWidth, type: WidthType.DXA },
                  alignment: AlignmentType.RIGHT,
                  layout: TableLayoutType.FIXED,
                  columnWidths: [badgeBoxWidth],
                  rows: [
                    new TableRow({
                      cantSplit: true,
                      children: [
                        simpleCell(badgeLabel, {
                          bold: true,
                          size: SZ_TITLE_WAIVER,
                          align: AlignmentType.CENTER,
                          widthTwips: badgeBoxWidth,
                          borders: ALL_THIN_BORDERS,
                          marginTB: 60,
                        }),
                      ],
                    }),
                  ],
                }),
              ],
              { widthTwips: headerCols[2], borders: ALL_NO_BORDERS, marginTB: 0, vAlign: VerticalAlign.TOP }
            ),
          ],
        }),
      ],
    })
  );

  const header = new Header({ children: headerChildren });

  const footer = new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 0, after: 0 },
        children: [
          txt('Page ', { size: SZ_BODY }),
          new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: SZ_BODY }),
          txt(' of ', { size: SZ_BODY }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: SZ_BODY }),
        ],
      }),
    ],
  });

  return { header, footer };
}

export async function generateICMemoDocxBlob(memo: ResolvedICMemo): Promise<Blob> {
  const { header, footer } = await buildBankHeaderAndFooter('Exit with Waiver');
  const body = buildICMemoCoreBlocks(memo, { includeClosingBlock: true });

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

export function buildICMemoCoreBlocks(
  memo: ResolvedICMemo,
  options: {
    includeClosingBlock?: boolean;
    metaOverride?: { dated?: string; memoNo?: string; to?: string; from?: string };
    skipMetadataTable?: boolean;
    insertedAfterSec15?: Array<{ sectionNumber: number; heading: string; text: string }>;
    particularsSectionNumber?: number;
    termsSectionNumber?: number;
  } = {}
): Array<Paragraph | Table> {
  const ph = memo.placeholders;
  const bp = memo.branchProposal;
  const v = (code: string) => valOrToken(ph, code);
  const insertedAfterSec15 = options.insertedAfterSec15 ?? [];
  const particularsSecNum = options.particularsSectionNumber ?? 16;
  const termsSecNum = options.termsSectionNumber ?? 17;

  // Document body children
  const body: Array<Paragraph | Table> = [];

  if (!options.skipMetadataTable) {
    // Top Metadata Table (Dated, Memo no, To, From) — ~44% of body width (4000 twips)
    const metaWidth = 4000;
    const metaCols = pctWidths([28, 7, 65], metaWidth);
    body.push(
      new Table({
        width: { size: metaWidth, type: WidthType.DXA },
        layout: TableLayoutType.FIXED,
        columnWidths: metaCols,
        rows: [
          new TableRow({
            cantSplit: true,
            children: [
              simpleCell('Dated', {
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
              simpleCell(options.metaOverride?.dated || v('#1#'), {
                bold: true,
                size: SZ_BODY,
                widthTwips: metaCols[2],
                borders: ALL_THIN_BORDERS,
                marginTB: 22,
              }),
            ],
          }),
          new TableRow({
            cantSplit: true,
            children: [
              simpleCell('Memo no', { bold: true, size: SZ_META, widthTwips: metaCols[0], borders: ALL_THIN_BORDERS, marginTB: 22 }),
              simpleCell(':', { bold: true, size: SZ_META, align: AlignmentType.CENTER, widthTwips: metaCols[1], borders: ALL_THIN_BORDERS, marginTB: 22 }),
              simpleCell(options.metaOverride?.memoNo ?? '', { size: SZ_BODY, bold: Boolean(options.metaOverride?.memoNo), widthTwips: metaCols[2], borders: ALL_THIN_BORDERS, marginTB: 22 }),
            ],
          }),
          new TableRow({
            cantSplit: true,
            children: [
              simpleCell('To', { bold: true, size: SZ_META, widthTwips: metaCols[0], borders: ALL_THIN_BORDERS, marginTB: 22 }),
              simpleCell(':', { bold: true, size: SZ_META, align: AlignmentType.CENTER, widthTwips: metaCols[1], borders: ALL_THIN_BORDERS, marginTB: 22 }),
              simpleCell(options.metaOverride?.to || 'Investment Committee', { bold: true, size: SZ_META, widthTwips: metaCols[2], borders: ALL_THIN_BORDERS, marginTB: 22 }),
            ],
          }),
          new TableRow({
            cantSplit: true,
            children: [
              simpleCell('From', { bold: true, size: SZ_META, widthTwips: metaCols[0], borders: ALL_THIN_BORDERS, marginTB: 22 }),
              simpleCell(':', { bold: true, size: SZ_META, align: AlignmentType.CENTER, widthTwips: metaCols[1], borders: ALL_THIN_BORDERS, marginTB: 22 }),
              simpleCell(options.metaOverride?.from || 'SAMD', { bold: true, size: SZ_META, widthTwips: metaCols[2], borders: ALL_THIN_BORDERS, marginTB: 22 }),
            ],
          }),
        ],
      }),
      p([], { before: 20, after: 20, line: 140 })
    );
  }

  // A/C and Branch Bar + Subject
  const acCols = pctWidths([58, 42]);
  body.push(
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: acCols,
      rows: [
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell(`A/C: ${v('#2#')} (CIS: ${v('#69#')})`, {
              bold: true,
              size: SZ_BODY,
              align: AlignmentType.CENTER,
              widthTwips: acCols[0],
              borders: ALL_DOTTED_BORDERS,
            }),
            simpleCell(`Branch: ${v('#3#')}`, {
              bold: true,
              size: SZ_BODY,
              align: AlignmentType.CENTER,
              widthTwips: acCols[1],
              borders: ALL_DOTTED_BORDERS,
            }),
          ],
        }),
        new TableRow({
          cantSplit: true,
          children: [
            cell(
              [
                p(
                  [
                    txt(
                      `Subject: The customer proposed for full and final settlement of their liability by allowing waiver of Tk.${v('#4#')} along with profit/compensation to be accrued up to the date of full and final settlement as follows:`,
                      { bold: true, size: SZ_BODY }
                    ),
                  ],
                  { before: 10, after: 10, line: 225 }
                ),
              ],
              { columnSpan: 2, widthTwips: BODY_WIDTH_TWIPS, borders: ALL_DOTTED_BORDERS }
            ),
          ],
        }),
      ],
    })
  );

  body.push(
    p(
      [
        txt('[All Figures in this Memo are in Million BDT Unless Stated Otherwise]', {
          bold: true,
          size: SZ_BODY,
        }),
      ],
      {
        align: AlignmentType.RIGHT,
        before: 30,
        after: 25,
      }
    )
  );

  // Main Settlement Table
  const mainCols = pctWidths([12, 8, 7, 21, 11, 11, 11, 19]);
  const remarksBlock = `1)Settlement Deadline:\n${v('#28#')}\n\n2) Break up\nPrincipal: ${v('#7#')}\nCharges: ${v('#8#')}\nCOF: ${v('#31#')}\nTotal: ${v('#32#')}\n\n3) Security Value:\nMV: ${v('#33#')}\nFSV: ${v('#34#')}`;

  body.push(
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: mainCols,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            simpleCell('Mode', { bold: true, rowSpan: 2, widthTwips: mainCols[0], align: AlignmentType.CENTER }),
            simpleCell('Limit', { bold: true, columnSpan: 2, widthTwips: mainCols[1] + mainCols[2], align: AlignmentType.CENTER }),
            simpleCell('Particulars', { bold: true, rowSpan: 2, widthTwips: mainCols[3], align: AlignmentType.CENTER }),
            simpleCell(`O/S as on ${v('#68#')}`, { bold: true, widthTwips: mainCols[4], align: AlignmentType.CENTER }),
            simpleCell('Proposed Waiver', { bold: true, widthTwips: mainCols[5], align: AlignmentType.CENTER }),
            simpleCell('Settlement Amount', { bold: true, widthTwips: mainCols[6], align: AlignmentType.CENTER }),
            simpleCell('Remarks', { bold: true, rowSpan: 2, widthTwips: mainCols[7], align: AlignmentType.CENTER }),
          ],
        }),
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            simpleCell('Original', { bold: true, align: AlignmentType.CENTER, widthTwips: mainCols[1] }),
            simpleCell('RSD', { bold: true, align: AlignmentType.CENTER, widthTwips: mainCols[2] }),
            simpleCell('01', { bold: true, align: AlignmentType.CENTER, widthTwips: mainCols[4] }),
            simpleCell('02', { bold: true, align: AlignmentType.CENTER, widthTwips: mainCols[5] }),
            simpleCell('03 (1-2)', { bold: true, align: AlignmentType.CENTER, widthTwips: mainCols[6] }),
          ],
        }),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell(v('#5#'), { bold: true, rowSpan: 7, align: AlignmentType.CENTER, widthTwips: mainCols[0] }),
            simpleCell(v('#6#'), { bold: true, rowSpan: 7, align: AlignmentType.CENTER, widthTwips: mainCols[1] }),
            simpleCell(bp?.limitRSD || '_', { rowSpan: 7, align: AlignmentType.CENTER, widthTwips: mainCols[2] }),
            simpleCell('Principal', { size: SZ_PARTICULARS, widthTwips: mainCols[3], marginTB: 22 }),
            simpleCell(v('#7#'), { bold: true, align: AlignmentType.CENTER, widthTwips: mainCols[4], marginTB: 22 }),
            simpleCell('', { align: AlignmentType.CENTER, widthTwips: mainCols[5], marginTB: 22 }),
            simpleCell('', { align: AlignmentType.CENTER, widthTwips: mainCols[6], marginTB: 22 }),
            simpleCell(remarksBlock, { bold: true, rowSpan: 7, vAlign: VerticalAlign.TOP, widthTwips: mainCols[7] }),
          ],
        }),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell('Charges (Legal & Others)', { size: SZ_PARTICULARS, widthTwips: mainCols[3], marginTB: 22 }),
            simpleCell(v('#8#'), { bold: true, align: AlignmentType.CENTER, widthTwips: mainCols[4], marginTB: 22 }),
            simpleCell('', { align: AlignmentType.CENTER, widthTwips: mainCols[5], marginTB: 22 }),
            simpleCell('', { align: AlignmentType.CENTER, widthTwips: mainCols[6], marginTB: 22 }),
          ],
        }),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell('Gestation Rent', { size: SZ_PARTICULARS, widthTwips: mainCols[3], marginTB: 22 }),
            simpleCell('-', { align: AlignmentType.CENTER, widthTwips: mainCols[4], marginTB: 22 }),
            simpleCell('', { align: AlignmentType.CENTER, widthTwips: mainCols[5], marginTB: 22 }),
            simpleCell('', { align: AlignmentType.CENTER, widthTwips: mainCols[6], marginTB: 22 }),
          ],
        }),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell('Charged Profit', { size: SZ_PARTICULARS, widthTwips: mainCols[3], marginTB: 22 }),
            simpleCell(v('#10#'), { bold: true, align: AlignmentType.CENTER, widthTwips: mainCols[4], marginTB: 22 }),
            simpleCell('', { align: AlignmentType.CENTER, widthTwips: mainCols[5], marginTB: 22 }),
            simpleCell('', { align: AlignmentType.CENTER, widthTwips: mainCols[6], marginTB: 22 }),
          ],
        }),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell('Compensation', { size: SZ_PARTICULARS, widthTwips: mainCols[3], marginTB: 22 }),
            simpleCell(v('#11#'), { bold: true, align: AlignmentType.CENTER, widthTwips: mainCols[4], marginTB: 22 }),
            simpleCell('', { align: AlignmentType.CENTER, widthTwips: mainCols[5], marginTB: 22 }),
            simpleCell('', { align: AlignmentType.CENTER, widthTwips: mainCols[6], marginTB: 22 }),
          ],
        }),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell('Uncharged Compensation', { size: SZ_PARTICULARS, widthTwips: mainCols[3], marginTB: 22 }),
            simpleCell(v('#12#'), { bold: true, align: AlignmentType.CENTER, widthTwips: mainCols[4], marginTB: 22 }),
            simpleCell('', { align: AlignmentType.CENTER, widthTwips: mainCols[5], marginTB: 22 }),
            simpleCell('', { align: AlignmentType.CENTER, widthTwips: mainCols[6], marginTB: 22 }),
          ],
        }),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell('Uncharged Profit (for T/L)', { size: SZ_PARTICULARS, widthTwips: mainCols[3], marginTB: 22 }),
            simpleCell('-', { align: AlignmentType.CENTER, widthTwips: mainCols[4], marginTB: 22 }),
            simpleCell('-', { align: AlignmentType.CENTER, widthTwips: mainCols[5], marginTB: 22 }),
            simpleCell('-', { align: AlignmentType.CENTER, widthTwips: mainCols[6], marginTB: 22 }),
          ],
        }),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell('Grand-Total', {
              bold: true,
              columnSpan: 4,
              align: AlignmentType.CENTER,
              widthTwips: mainCols[0] + mainCols[1] + mainCols[2] + mainCols[3],
              marginTB: 25,
            }),
            simpleCell(v('#25#'), { bold: true, align: AlignmentType.CENTER, widthTwips: mainCols[4], marginTB: 25 }),
            simpleCell(v('#4#'), { bold: true, align: AlignmentType.CENTER, widthTwips: mainCols[5], marginTB: 25 }),
            simpleCell(v('#27#'), { bold: true, align: AlignmentType.CENTER, widthTwips: mainCols[6], marginTB: 25 }),
            simpleCell('', { widthTwips: mainCols[7], marginTB: 25 }),
          ],
        }),
      ],
    })
  );

  // Note line
  body.push(
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: [BODY_WIDTH_TWIPS],
      rows: [
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell(
              `Note: Information Regarding Down payment (deposited Tk.${v('#35#')} Million in A/C: ${v('#36#')} on ${v('#68#')}.)`,
              { bold: true, size: SZ_BODY, widthTwips: BODY_WIDTH_TWIPS, marginTB: 25 }
            ),
          ],
        }),
      ],
    })
  );

  // Key Points
  const kpCols = pctWidths([5, 42, 11, 5, 22, 15]);
  body.push(
    p([txt('Key Points:', { bold: true, underline: true })], { before: 45, after: 15 }),
    p(
      [
        txt('This Memo is placed in conformity with ', { size: SZ_BODY }),
        txt('BRPD 1 Circular Letter No. 23 dated 29.06.2026', { bold: true, size: SZ_BODY }),
        txt(' under the following contexts:', { size: SZ_BODY }),
      ],
      { before: 15, after: 25 }
    ),
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: kpCols,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            simpleCell('Particulars', {
              bold: true,
              columnSpan: 3,
              widthTwips: kpCols[0] + kpCols[1] + kpCols[2],
              marginTB: 20,
            }),
            simpleCell('Waiver recommended by:', {
              bold: true,
              columnSpan: 3,
              widthTwips: kpCols[3] + kpCols[4] + kpCols[5],
              marginTB: 20,
            }),
          ],
        }),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell('a.', { widthTwips: kpCols[0], marginTB: 20 }),
            simpleCell('Total [Profit+ Compensation+ Uncharged Comp.& Profit]', { widthTwips: kpCols[1], marginTB: 20 }),
            simpleCell(`: ${v('#42#')}`, { bold: true, widthTwips: kpCols[2], marginTB: 20 }),
            simpleCell('d.', { widthTwips: kpCols[3], marginTB: 20 }),
            simpleCell('Client', { widthTwips: kpCols[4], marginTB: 20 }),
            simpleCell(`: ${v('#45#')}`, { bold: true, widthTwips: kpCols[5], marginTB: 20 }),
          ],
        }),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell('b.', { widthTwips: kpCols[0], marginTB: 20 }),
            simpleCell('Cost of Fund', { widthTwips: kpCols[1], marginTB: 20 }),
            simpleCell(`: ${v('#31#')}`, { bold: true, widthTwips: kpCols[2], marginTB: 20 }),
            simpleCell('e.', { widthTwips: kpCols[3], marginTB: 20 }),
            simpleCell('Branch', { widthTwips: kpCols[4], marginTB: 20 }),
            simpleCell(`: ${v('#45#')}`, { bold: true, widthTwips: kpCols[5], marginTB: 20 }),
          ],
        }),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell('c.', { rowSpan: 4, widthTwips: kpCols[0], marginTB: 20 }),
            simpleCell('Maximum Waiver Considering COF [a-b]', { rowSpan: 4, widthTwips: kpCols[1], marginTB: 20 }),
            simpleCell(`: ${v('#44#')}`, { bold: true, rowSpan: 4, widthTwips: kpCols[2], marginTB: 20 }),
            simpleCell('f.', { widthTwips: kpCols[3], marginTB: 20 }),
            simpleCell('SAMD, HO', { widthTwips: kpCols[4], marginTB: 20 }),
            simpleCell(`: ${v('#47#')}`, { bold: true, widthTwips: kpCols[5], marginTB: 20 }),
          ],
        }),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell('g.', { widthTwips: kpCols[3], marginTB: 20 }),
            simpleCell('HOIC', { widthTwips: kpCols[4], marginTB: 20 }),
            simpleCell(`: ${v('#47#')}`, { bold: true, widthTwips: kpCols[5], marginTB: 20 }),
          ],
        }),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell('h.', { widthTwips: kpCols[3], marginTB: 20 }),
            simpleCell('MEC', { widthTwips: kpCols[4], marginTB: 20 }),
            simpleCell(`: ${v('#47#')}`, { bold: true, widthTwips: kpCols[5], marginTB: 20 }),
          ],
        }),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell('i.', { widthTwips: kpCols[3], marginTB: 20 }),
            simpleCell('IC & CD', { widthTwips: kpCols[4], marginTB: 20 }),
            simpleCell(`: ${v('#47#')}`, { bold: true, widthTwips: kpCols[5], marginTB: 20 }),
          ],
        }),
      ],
    })
  );

  // Justification of The Proposed Facility
  const justLines = memo.justificationLines
    ? memo.justificationLines.split('\n').map((l) => l.trim()).filter(Boolean)
    : [];
  body.push(
    p([txt('Justification of The Proposed Facility:', { bold: true, underline: true })], { before: 45, after: 15 })
  );
  if (justLines.length > 1) {
    body.push(p([txt('a. ', { bold: true }), txt(justLines[0])], { indentLeft: 120, before: 15, after: 10 }));
    for (let idx = 1; idx < justLines.length; idx++) {
      body.push(p([txt(justLines[idx])], { indentLeft: 280, before: 10, after: 10 }));
    }
  } else if (justLines.length === 1) {
    body.push(p([txt('a. ', { bold: true }), txt(justLines[0])], { indentLeft: 120, before: 15, after: 15 }));
  } else {
    body.push(p([txt(`a. ${v('#51#')}`)], { indentLeft: 120, before: 15, after: 15 }));
  }
  body.push(
    p([txt('b. Investment was not created against any fraud and forgery.')], { indentLeft: 120, before: 15, after: 15 }),
    p([txt('c. The client is not a willful defaulter.')], { indentLeft: 120, before: 15, after: 15 }),
    p([txt('Previous Board Decision(s) Regarding Rescheduling/Waiver/Written-Off: Nil', { bold: true })], {
      before: 30,
      after: 30,
    })
  );

  // Section 1
  body.push(
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: [BODY_WIDTH_TWIPS],
      rows: [
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell(`1. Nature of Business: ${v('#54#')}`, {
              bold: true,
              size: SZ_BODY,
              widthTwips: BODY_WIDTH_TWIPS,
              marginTB: 22,
            }),
          ],
        }),
      ],
    })
  );

  // Section 2 (matching IC_Memo.pdf Page 1 with 2 spacer rows between Funded and Non Funded)
  const sec2Cols = pctWidths([14, 17, 10, 11, 11, 10, 17, 10]);
  body.push(
    p(
      [
        txt(
          `2. Liability Position of the Clients and Its Sister Concerns with Our Bank as on ${v('#68#')}:`,
          { bold: true }
        ),
      ],
      { before: 40, after: 20 }
    ),
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: sec2Cols,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            simpleCell('Name of the Concern', { bold: true, widthTwips: sec2Cols[0], align: AlignmentType.LEFT, marginTB: 22 }),
            simpleCell('Mode of Investment', { bold: true, widthTwips: sec2Cols[1], align: AlignmentType.LEFT, marginTB: 22 }),
            simpleCell('Limit', { bold: true, widthTwips: sec2Cols[2], align: AlignmentType.CENTER, marginTB: 22 }),
            simpleCell('Outstanding', { bold: true, widthTwips: sec2Cols[3], align: AlignmentType.CENTER, marginTB: 22 }),
            simpleCell('Validity', { bold: true, widthTwips: sec2Cols[4], align: AlignmentType.CENTER, marginTB: 22 }),
            simpleCell('Overdue', { bold: true, widthTwips: sec2Cols[5], align: AlignmentType.CENTER, marginTB: 22 }),
            simpleCell('Recovery in existing deals till date [Mode wise]', {
              bold: true,
              widthTwips: sec2Cols[6],
              align: AlignmentType.CENTER,
              marginTB: 22,
            }),
            simpleCell('Remarks', { bold: true, widthTwips: sec2Cols[7], align: AlignmentType.LEFT, marginTB: 22 }),
          ],
        }),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell(v('#2#'), { bold: true, rowSpan: 4, widthTwips: sec2Cols[0], marginTB: 20 }),
            simpleCell(`${v('#55#')}\nFunded`, { bold: true, widthTwips: sec2Cols[1], marginTB: 20 }),
            simpleCell(v('#56#'), { bold: true, align: AlignmentType.CENTER, widthTwips: sec2Cols[2], marginTB: 20 }),
            simpleCell(v('#25#'), { bold: true, align: AlignmentType.CENTER, widthTwips: sec2Cols[3], marginTB: 20 }),
            simpleCell(v('#58#'), { bold: true, align: AlignmentType.CENTER, widthTwips: sec2Cols[4], marginTB: 20 }),
            simpleCell(v('#59#'), { bold: true, align: AlignmentType.CENTER, widthTwips: sec2Cols[5], marginTB: 20 }),
            simpleCell(v('#60#'), { bold: true, align: AlignmentType.CENTER, widthTwips: sec2Cols[6], marginTB: 20 }),
            simpleCell(bp?.sec2Remarks || '', { widthTwips: sec2Cols[7], marginTB: 20 }),
          ],
        }),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell('', { widthTwips: sec2Cols[1], marginTB: 20 }),
            simpleCell('', { widthTwips: sec2Cols[2], marginTB: 20 }),
            simpleCell('', { widthTwips: sec2Cols[3], marginTB: 20 }),
            simpleCell('', { widthTwips: sec2Cols[4], marginTB: 20 }),
            simpleCell('', { widthTwips: sec2Cols[5], marginTB: 20 }),
            simpleCell('', { widthTwips: sec2Cols[6], marginTB: 20 }),
            simpleCell('', { widthTwips: sec2Cols[7], marginTB: 20 }),
          ],
        }),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell('', { widthTwips: sec2Cols[1], marginTB: 20 }),
            simpleCell('', { widthTwips: sec2Cols[2], marginTB: 20 }),
            simpleCell('', { widthTwips: sec2Cols[3], marginTB: 20 }),
            simpleCell('', { widthTwips: sec2Cols[4], marginTB: 20 }),
            simpleCell('', { widthTwips: sec2Cols[5], marginTB: 20 }),
            simpleCell('', { widthTwips: sec2Cols[6], marginTB: 20 }),
            simpleCell('', { widthTwips: sec2Cols[7], marginTB: 20 }),
          ],
        }),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell('Non Funded', { bold: true, widthTwips: sec2Cols[1], marginTB: 20 }),
            simpleCell('', { widthTwips: sec2Cols[2], marginTB: 20 }),
            simpleCell('', { widthTwips: sec2Cols[3], marginTB: 20 }),
            simpleCell('', { widthTwips: sec2Cols[4], marginTB: 20 }),
            simpleCell('', { widthTwips: sec2Cols[5], marginTB: 20 }),
            simpleCell('', { widthTwips: sec2Cols[6], marginTB: 20 }),
            simpleCell(memo.sec2NonFundedNote || '', { widthTwips: sec2Cols[7], marginTB: 20 }),
          ],
        }),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell('Total', {
              bold: true,
              columnSpan: 2,
              align: AlignmentType.CENTER,
              widthTwips: sec2Cols[0] + sec2Cols[1],
              marginTB: 20,
            }),
            simpleCell(v('#56#'), { bold: true, align: AlignmentType.CENTER, widthTwips: sec2Cols[2], marginTB: 20 }),
            simpleCell(v('#25#'), { bold: true, align: AlignmentType.CENTER, widthTwips: sec2Cols[3], marginTB: 20 }),
            simpleCell(v('#58#'), { bold: true, align: AlignmentType.CENTER, widthTwips: sec2Cols[4], marginTB: 20 }),
            simpleCell(v('#59#'), { bold: true, align: AlignmentType.CENTER, widthTwips: sec2Cols[5], marginTB: 20 }),
            simpleCell(v('#60#'), { bold: true, align: AlignmentType.CENTER, widthTwips: sec2Cols[6], marginTB: 20 }),
            simpleCell('', { widthTwips: sec2Cols[7], marginTB: 20 }),
          ],
        }),
      ],
    })
  );

  // Section 3 (Starts at top of Page 2 in exact reference IC_Memo.pdf)
  const rawSec3Rows =
    bp?.section3Rows && bp.section3Rows.length > 0
      ? bp.section3Rows
      : [
          { sl: 'a.', nameOfConcern: '', bankOrNbfi: '', natureOfFacility: '', limit: '', outstanding: '', overdueEol: '', expiryDate: '', clStatus: '' },
          { sl: 'b.', nameOfConcern: '', bankOrNbfi: '', natureOfFacility: '', limit: '', outstanding: '', overdueEol: '', expiryDate: '', clStatus: '' },
          { sl: 'c.', nameOfConcern: '', bankOrNbfi: '', natureOfFacility: '', limit: '', outstanding: '', overdueEol: '', expiryDate: '', clStatus: '' },
        ];
  const hasSec3SubTotal = rawSec3Rows.some((r) => /sub[-\s]*total/i.test(r.sl) || /sub[-\s]*total/i.test(r.nameOfConcern));
  const hasSec3GrandTotal = rawSec3Rows.some((r) => /grand\s*total/i.test(r.sl) || /grand\s*total/i.test(r.nameOfConcern));
  const sec3Cols = pctWidths([5, 13, 17, 11, 8, 17, 11, 10, 8]);

  body.push(
    p(
      [
        txt('3. Liability of the Client & Its Sister/Allied Concerns with Other Units of SIB:', {
          bold: true,
          underline: true,
        }),
      ],
      { before: 0, after: 20, pageBreakBefore: true }
    ),
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: sec3Cols,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            simpleCell('SL', { bold: true, widthTwips: sec3Cols[0], align: AlignmentType.CENTER, marginTB: 22 }),
            simpleCell('Name of Concern', { bold: true, widthTwips: sec3Cols[1], marginTB: 22 }),
            simpleCell('Name of the Banks/NBFIs', { bold: true, widthTwips: sec3Cols[2], marginTB: 22 }),
            simpleCell('Nature of Facility', { bold: true, widthTwips: sec3Cols[3], marginTB: 22 }),
            simpleCell('Limit', { bold: true, widthTwips: sec3Cols[4], align: AlignmentType.CENTER, marginTB: 22 }),
            simpleCell('Outstanding with profit/Interest', { bold: true, widthTwips: sec3Cols[5], align: AlignmentType.CENTER, marginTB: 22 }),
            simpleCell('Overdue/EOL', { bold: true, widthTwips: sec3Cols[6], align: AlignmentType.CENTER, marginTB: 22 }),
            simpleCell('Expiry Date', { bold: true, widthTwips: sec3Cols[7], align: AlignmentType.CENTER, marginTB: 22 }),
            simpleCell('CL Status', { bold: true, widthTwips: sec3Cols[8], align: AlignmentType.CENTER, marginTB: 22 }),
          ],
        }),
        ...rawSec3Rows.map((r) => {
          const isTotalRow = /sub[-\s]*total|grand\s*total/i.test(r.sl) || /sub[-\s]*total|grand\s*total/i.test(r.nameOfConcern);
          return new TableRow({
            cantSplit: true,
            children: [
              simpleCell(r.sl, { bold: isTotalRow, align: AlignmentType.CENTER, widthTwips: sec3Cols[0], marginTB: 20 }),
              simpleCell(r.nameOfConcern, { bold: isTotalRow, widthTwips: sec3Cols[1], marginTB: 20 }),
              simpleCell(r.bankOrNbfi, { bold: isTotalRow, widthTwips: sec3Cols[2], marginTB: 20 }),
              simpleCell(r.natureOfFacility, { bold: isTotalRow, widthTwips: sec3Cols[3], marginTB: 20 }),
              simpleCell(r.limit, { bold: isTotalRow, align: AlignmentType.CENTER, widthTwips: sec3Cols[4], marginTB: 20 }),
              simpleCell(r.outstanding, { bold: isTotalRow, align: AlignmentType.CENTER, widthTwips: sec3Cols[5], marginTB: 20 }),
              simpleCell(r.overdueEol, { bold: isTotalRow, align: AlignmentType.CENTER, widthTwips: sec3Cols[6], marginTB: 20 }),
              simpleCell(r.expiryDate, { bold: isTotalRow, align: AlignmentType.CENTER, widthTwips: sec3Cols[7], marginTB: 20 }),
              simpleCell(r.clStatus, { bold: isTotalRow, align: AlignmentType.CENTER, widthTwips: sec3Cols[8], marginTB: 20 }),
            ],
          });
        }),
        ...(!hasSec3SubTotal
          ? [
              new TableRow({
                cantSplit: true,
                children: [
                  simpleCell('Sub-Total', {
                    bold: true,
                    columnSpan: 4,
                    align: AlignmentType.CENTER,
                    widthTwips: sec3Cols[0] + sec3Cols[1] + sec3Cols[2] + sec3Cols[3],
                    marginTB: 20,
                  }),
                  simpleCell('', { widthTwips: sec3Cols[4], marginTB: 20 }),
                  simpleCell('', { widthTwips: sec3Cols[5], marginTB: 20 }),
                  simpleCell('', { widthTwips: sec3Cols[6], marginTB: 20 }),
                  simpleCell('', { widthTwips: sec3Cols[7], marginTB: 20 }),
                  simpleCell('', { widthTwips: sec3Cols[8], marginTB: 20 }),
                ],
              }),
            ]
          : []),
        ...(!hasSec3GrandTotal
          ? [
              new TableRow({
                cantSplit: true,
                children: [
                  simpleCell('Grand Total', {
                    bold: true,
                    columnSpan: 4,
                    align: AlignmentType.CENTER,
                    widthTwips: sec3Cols[0] + sec3Cols[1] + sec3Cols[2] + sec3Cols[3],
                    marginTB: 20,
                  }),
                  simpleCell('', { widthTwips: sec3Cols[4], marginTB: 20 }),
                  simpleCell('', { widthTwips: sec3Cols[5], marginTB: 20 }),
                  simpleCell('', { widthTwips: sec3Cols[6], marginTB: 20 }),
                  simpleCell('', { widthTwips: sec3Cols[7], marginTB: 20 }),
                  simpleCell('', { widthTwips: sec3Cols[8], marginTB: 20 }),
                ],
              }),
            ]
          : []),
      ],
    })
  );

  // Section 4 (CIB Status)
  const defaultCibStatuses = [
    'STD',
    'SMA',
    'SS',
    'DF',
    'BL',
    'Willful Default (WD)',
    'Write-Off',
    'Non-Funded',
    'Stay Order',
    'TOTAL',
  ];
  const sec4Rows =
    bp?.section4Rows && bp.section4Rows.length > 0
      ? bp.section4Rows
      : defaultCibStatuses.map((s) => ({
          status: s,
          unitBorrower: '',
          unitGuarantor: '',
          otherBorrower: '',
          otherGuarantor: '',
          remarks: '',
        }));
  const sec4Cols = pctWidths([18, 13, 13, 13, 13, 30]);

  body.push(
    p(
      [
        txt('4. CIB Status: ', { bold: true, underline: true }),
        txt(`Report dated ${v('#62#')} shows:`, { bold: true }),
      ],
      { before: 50, after: 20 }
    ),
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: sec4Cols,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            simpleCell('STATUS', { bold: true, rowSpan: 2, widthTwips: sec4Cols[0], marginTB: 22 }),
            simpleCell('With respective unit', {
              bold: true,
              columnSpan: 2,
              widthTwips: sec4Cols[1] + sec4Cols[2],
              align: AlignmentType.CENTER,
              marginTB: 22,
            }),
            simpleCell('With Other Bank', {
              bold: true,
              columnSpan: 2,
              widthTwips: sec4Cols[3] + sec4Cols[4],
              align: AlignmentType.CENTER,
              marginTB: 22,
            }),
            simpleCell('Remarks\n(Pls. give concern wise Breakup)', {
              bold: true,
              rowSpan: 2,
              widthTwips: sec4Cols[5],
              align: AlignmentType.CENTER,
              marginTB: 22,
            }),
          ],
        }),
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            simpleCell('Borrower', { bold: true, align: AlignmentType.CENTER, widthTwips: sec4Cols[1], marginTB: 20 }),
            simpleCell('Guarantor', { bold: true, align: AlignmentType.CENTER, widthTwips: sec4Cols[2], marginTB: 20 }),
            simpleCell('Borrower', { bold: true, align: AlignmentType.CENTER, widthTwips: sec4Cols[3], marginTB: 20 }),
            simpleCell('Guarantor', { bold: true, align: AlignmentType.CENTER, widthTwips: sec4Cols[4], marginTB: 20 }),
          ],
        }),
        ...sec4Rows.map(
          (r) =>
            new TableRow({
              cantSplit: true,
              children: [
                simpleCell(r.status, { bold: /TOTAL/i.test(r.status), widthTwips: sec4Cols[0], marginTB: 20 }),
                simpleCell(r.unitBorrower, { align: AlignmentType.CENTER, bold: /TOTAL/i.test(r.status), widthTwips: sec4Cols[1], marginTB: 20 }),
                simpleCell(r.unitGuarantor, { align: AlignmentType.CENTER, widthTwips: sec4Cols[2], marginTB: 20 }),
                simpleCell(r.otherBorrower, { align: AlignmentType.CENTER, widthTwips: sec4Cols[3], marginTB: 20 }),
                simpleCell(r.otherGuarantor, { align: AlignmentType.CENTER, widthTwips: sec4Cols[4], marginTB: 20 }),
                simpleCell(r.remarks, { widthTwips: sec4Cols[5], marginTB: 20 }),
              ],
            })
        ),
      ],
    })
  );

  // Section 5 (Legal Action Status)
  const sec5Rows =
    bp?.section5Rows && bp.section5Rows.length > 0
      ? bp.section5Rows
      : [{ sl: '', typeOfCase: '', caseNo: '', filingDate: '', amount: '', previousDate: '', nextDate: '', presentStatus: '' }];
  const sec5Cols = pctWidths([5, 16, 13, 12, 12, 12, 12, 18]);

  body.push(
    p(
      [
        txt('5. Legal Action Status as on Memo Date:', { bold: true, underline: true }),
        ...(bp?.section5StatusNote ? [txt(` ${bp.section5StatusNote}`, { bold: true })] : []),
      ],
      { before: 50, after: 20 }
    ),
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: sec5Cols,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            simpleCell('Sl', { bold: true, align: AlignmentType.CENTER, widthTwips: sec5Cols[0], marginTB: 22 }),
            simpleCell('Type of Case', { bold: true, widthTwips: sec5Cols[1], marginTB: 22 }),
            simpleCell('Case No.', { bold: true, widthTwips: sec5Cols[2], marginTB: 22 }),
            simpleCell('Filling date', { bold: true, align: AlignmentType.CENTER, widthTwips: sec5Cols[3], marginTB: 22 }),
            simpleCell('Amount', { bold: true, align: AlignmentType.CENTER, widthTwips: sec5Cols[4], marginTB: 22 }),
            simpleCell('Previous date', { bold: true, align: AlignmentType.CENTER, widthTwips: sec5Cols[5], marginTB: 22 }),
            simpleCell('Next date', { bold: true, align: AlignmentType.CENTER, widthTwips: sec5Cols[6], marginTB: 22 }),
            simpleCell('Present status', { bold: true, widthTwips: sec5Cols[7], marginTB: 22 }),
          ],
        }),
        ...sec5Rows.map((r) => {
          const isEmpty = !r.sl && !r.typeOfCase && !r.caseNo;
          const pad = isEmpty ? 120 : 25;
          return new TableRow({
            cantSplit: true,
            children: [
              simpleCell(r.sl, { align: AlignmentType.CENTER, widthTwips: sec5Cols[0], marginTB: pad }),
              simpleCell(r.typeOfCase, { widthTwips: sec5Cols[1], marginTB: pad }),
              simpleCell(r.caseNo, { widthTwips: sec5Cols[2], marginTB: pad }),
              simpleCell(r.filingDate, { align: AlignmentType.CENTER, widthTwips: sec5Cols[3], marginTB: pad }),
              simpleCell(r.amount, { align: AlignmentType.CENTER, widthTwips: sec5Cols[4], marginTB: pad }),
              simpleCell(r.previousDate, { align: AlignmentType.CENTER, widthTwips: sec5Cols[5], marginTB: pad }),
              simpleCell(r.nextDate, { align: AlignmentType.CENTER, widthTwips: sec5Cols[6], marginTB: pad }),
              simpleCell(r.presentStatus, { widthTwips: sec5Cols[7], marginTB: pad }),
            ],
          });
        }),
      ],
    })
  );

  // Section 6 (Audit, Compliance)
  const defaultSec6 = [
    { sl: 'a.', authority: 'IC & CD', lastInspectionDate: '', objectionStatus: '', objections: '', branchReply: '' },
    { sl: 'b.', authority: 'DBI, Bangladesh Bank', lastInspectionDate: '', objectionStatus: 'No', objections: '', branchReply: '' },
    { sl: 'c.', authority: 'DOS, Bangladesh Bank', lastInspectionDate: 'Not done', objectionStatus: 'No', objections: '', branchReply: '' },
    { sl: 'd.', authority: 'BFIU, Bangladesh Bank', lastInspectionDate: 'Not done', objectionStatus: 'No', objections: '', branchReply: '' },
    { sl: 'e', authority: 'BRD, Bangladesh Bank', lastInspectionDate: 'Not done', objectionStatus: 'No', objections: '', branchReply: '' },
    { sl: 'f', authority: 'ACC(Anti-Corruption Commission)', lastInspectionDate: 'Not done', objectionStatus: 'No', objections: '', branchReply: '' },
  ];
  const sec6Rows = bp?.section6Rows && bp.section6Rows.length > 0 ? bp.section6Rows : defaultSec6;
  const sec6Cols = pctWidths([5, 24, 18, 15, 19, 19]);

  body.push(
    p(
      [
        txt('6. Audit, Compliance and Regulatory Authority Related Status:', { bold: true, underline: true }),
        ...(bp?.section6StatusNote ? [txt(` ${bp.section6StatusNote}`, { bold: true })] : []),
      ],
      { before: 50, after: 20 }
    ),
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: sec6Cols,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            simpleCell('SL', { bold: true, align: AlignmentType.CENTER, widthTwips: sec6Cols[0], marginTB: 22 }),
            simpleCell('Name of the Authority', { bold: true, widthTwips: sec6Cols[1], marginTB: 22 }),
            simpleCell('Last Inspection / Report Date', { bold: true, align: AlignmentType.CENTER, widthTwips: sec6Cols[2], marginTB: 22 }),
            simpleCell('Objection Status (Yes/No)', { bold: true, align: AlignmentType.CENTER, widthTwips: sec6Cols[3], marginTB: 22 }),
            simpleCell('Objections (If yes)', { bold: true, widthTwips: sec6Cols[4], marginTB: 22 }),
            simpleCell('Branch Reply/ Current Position', { bold: true, widthTwips: sec6Cols[5], marginTB: 22 }),
          ],
        }),
        ...sec6Rows.map(
          (r) =>
            new TableRow({
              cantSplit: true,
              children: [
                simpleCell(r.sl, { align: AlignmentType.CENTER, widthTwips: sec6Cols[0], marginTB: 20 }),
                simpleCell(r.authority, { widthTwips: sec6Cols[1], marginTB: 20 }),
                simpleCell(r.lastInspectionDate, { align: AlignmentType.CENTER, widthTwips: sec6Cols[2], marginTB: 20 }),
                simpleCell(r.objectionStatus, { align: AlignmentType.CENTER, widthTwips: sec6Cols[3], marginTB: 20 }),
                simpleCell(r.objections, { widthTwips: sec6Cols[4], marginTB: 20 }),
                simpleCell(r.branchReply, { widthTwips: sec6Cols[5], marginTB: 20 }),
              ],
            })
        ),
      ],
    })
  );

  // Section 7 (Security)
  const defaultSec7Primary = [
    { particulars: 'i) Hypothecation of Stock as on ………………..', value: '' },
    { particulars: 'ii) Hypothecation of Machinery', value: '' },
    { particulars: 'iii) Assignment of Bills against Work Order', value: '' },
    { particulars: 'iv) Margin held against existing LC & BG liability as on……..', value: '' },
  ];
  const sec7Primary =
    bp?.section7PrimaryRows && bp.section7PrimaryRows.length > 0 ? bp.section7PrimaryRows : defaultSec7Primary;
  const defaultSchedules = [
    'Schedule-01: Full Schedule with owner name as per Mortgage Deed.',
    'Schedule-02: Full Schedule with owner name as per Mortgage Deed.',
    'Schedule-03: Full Schedule with owner name as per Mortgage Deed.',
    'Schedule-04: Full Schedule with owner name as per Mortgage Deed.',
  ];
  const schedules =
    bp?.section7CollateralSchedules && bp.section7CollateralSchedules.length > 0
      ? bp.section7CollateralSchedules
      : defaultSchedules;

  // Exact order in IC_Memo.pdf Page 3: Schedule-01, Schedule-03, Schedule-02, Schedule-04, Building, Other, Grand Total
  const defaultValRows = [
    'Schedule-01',
    'Schedule-03',
    'Schedule-02',
    'Schedule-04',
    'Building',
    'Other',
    'Grand Total',
  ].map((s) => ({
    schedule: s,
    origBranchMv: '',
    origBranchFsv: '',
    origSurvMv: '',
    origSurvFsv: '',
    presBranchMv: '',
    presBranchFsv: '',
    presSurvMv: '',
    presSurvFsv: '',
    presAvMv: '',
    presAvFsv: '',
  }));
  const valRows =
    bp?.section7ValuationRows && bp.section7ValuationRows.length > 0 ? bp.section7ValuationRows : defaultValRows;

  const sec7PrimCols = pctWidths([75, 25]);
  const sec7ValCols = pctWidths([13, 8.7, 8.7, 8.7, 8.7, 8.7, 8.7, 8.7, 8.7, 8.7, 8.7]);

  body.push(
    p([txt('7. Security:', { bold: true, underline: true })], { before: 50, after: 15 }),
    p(
      [
        txt('a. Primary:', { bold: true, underline: true }),
        ...(bp?.section7PrimaryNote ? [txt(` ${bp.section7PrimaryNote}`, { bold: true })] : []),
      ],
      { indentLeft: 120, before: 15, after: 15 }
    ),
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: sec7PrimCols,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            simpleCell('Particulars', { bold: true, widthTwips: sec7PrimCols[0], marginTB: 22 }),
            simpleCell('Value', { bold: true, widthTwips: sec7PrimCols[1], align: AlignmentType.CENTER, marginTB: 22 }),
          ],
        }),
        ...sec7Primary.map(
          (r) =>
            new TableRow({
              cantSplit: true,
              children: [
                simpleCell(r.particulars, { widthTwips: sec7PrimCols[0], marginTB: 20 }),
                simpleCell(r.value, { align: AlignmentType.CENTER, widthTwips: sec7PrimCols[1], marginTB: 20 }),
              ],
            })
        ),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell('Total value of primary security', {
              bold: true,
              align: AlignmentType.CENTER,
              widthTwips: sec7PrimCols[0],
              marginTB: 22,
            }),
            simpleCell(bp?.section7PrimaryTotal || '', {
              bold: true,
              align: AlignmentType.CENTER,
              widthTwips: sec7PrimCols[1],
              marginTB: 22,
            }),
          ],
        }),
      ],
    }),
    p(
      [
        txt('b. Collateral:', { bold: true, underline: true }),
        ...(bp?.section7CollateralNote ? [txt(` ${bp.section7CollateralNote}`, { bold: true })] : []),
      ],
      { indentLeft: 120, before: 35, after: 15 }
    ),
    ...schedules.map((s) => p([txt(s, { size: SZ_BODY })], { indentLeft: 120, before: 15, after: 15 })),
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: sec7ValCols,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            simpleCell('Schedule', {
              bold: true,
              size: SZ_SMALL,
              rowSpan: 4,
              widthTwips: sec7ValCols[0],
              vAlign: VerticalAlign.CENTER,
              marginTB: 20,
            }),
            simpleCell('Value of the mortgaged properties as assessed', {
              bold: true,
              size: SZ_SMALL,
              columnSpan: 10,
              widthTwips: BODY_WIDTH_TWIPS - sec7ValCols[0],
              align: AlignmentType.CENTER,
              marginTB: 20,
            }),
          ],
        }),
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            simpleCell('At the time of allowing investment', {
              bold: true,
              size: SZ_SMALL,
              columnSpan: 4,
              widthTwips: sec7ValCols[1] + sec7ValCols[2] + sec7ValCols[3] + sec7ValCols[4],
              align: AlignmentType.CENTER,
              marginTB: 18,
            }),
            simpleCell('At present', {
              bold: true,
              size: SZ_SMALL,
              columnSpan: 6,
              widthTwips: sec7ValCols[5] + sec7ValCols[6] + sec7ValCols[7] + sec7ValCols[8] + sec7ValCols[9] + sec7ValCols[10],
              align: AlignmentType.CENTER,
              marginTB: 18,
            }),
          ],
        }),
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            simpleCell('By Branch\non………..', { bold: true, size: SZ_SMALL, columnSpan: 2, widthTwips: sec7ValCols[1] + sec7ValCols[2], align: AlignmentType.CENTER, marginTB: 16 }),
            simpleCell('Surveyor\non……………..', { bold: true, size: SZ_SMALL, columnSpan: 2, widthTwips: sec7ValCols[3] + sec7ValCols[4], align: AlignmentType.CENTER, marginTB: 16 }),
            simpleCell('By Branch\non……………. ', { bold: true, size: SZ_SMALL, columnSpan: 2, widthTwips: sec7ValCols[5] + sec7ValCols[6], align: AlignmentType.CENTER, marginTB: 16 }),
            simpleCell('Surveyor\non………….', { bold: true, size: SZ_SMALL, columnSpan: 2, widthTwips: sec7ValCols[7] + sec7ValCols[8], align: AlignmentType.CENTER, marginTB: 16 }),
            simpleCell('By AV Cell\non……………', { bold: true, size: SZ_SMALL, columnSpan: 2, widthTwips: sec7ValCols[9] + sec7ValCols[10], align: AlignmentType.CENTER, marginTB: 16 }),
          ],
        }),
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            simpleCell('MV', { bold: true, size: SZ_SMALL, align: AlignmentType.CENTER, widthTwips: sec7ValCols[1], marginTB: 16 }),
            simpleCell('FSV', { bold: true, size: SZ_SMALL, align: AlignmentType.CENTER, widthTwips: sec7ValCols[2], marginTB: 16 }),
            simpleCell('MV', { bold: true, size: SZ_SMALL, align: AlignmentType.CENTER, widthTwips: sec7ValCols[3], marginTB: 16 }),
            simpleCell('FSV', { bold: true, size: SZ_SMALL, align: AlignmentType.CENTER, widthTwips: sec7ValCols[4], marginTB: 16 }),
            simpleCell('MV', { bold: true, size: SZ_SMALL, align: AlignmentType.CENTER, widthTwips: sec7ValCols[5], marginTB: 16 }),
            simpleCell('FSV', { bold: true, size: SZ_SMALL, align: AlignmentType.CENTER, widthTwips: sec7ValCols[6], marginTB: 16 }),
            simpleCell('MV', { bold: true, size: SZ_SMALL, align: AlignmentType.CENTER, widthTwips: sec7ValCols[7], marginTB: 16 }),
            simpleCell('FSV', { bold: true, size: SZ_SMALL, align: AlignmentType.CENTER, widthTwips: sec7ValCols[8], marginTB: 16 }),
            simpleCell('MV', { bold: true, size: SZ_SMALL, align: AlignmentType.CENTER, widthTwips: sec7ValCols[9], marginTB: 16 }),
            simpleCell('FSV', { bold: true, size: SZ_SMALL, align: AlignmentType.CENTER, widthTwips: sec7ValCols[10], marginTB: 16 }),
          ],
        }),
        ...valRows.map(
          (r) =>
            new TableRow({
              cantSplit: true,
              children: [
                simpleCell(r.schedule, {
                  size: SZ_SMALL,
                  bold: /Grand/i.test(r.schedule),
                  widthTwips: sec7ValCols[0],
                  marginTB: 18,
                }),
                simpleCell(r.origBranchMv, { size: SZ_SMALL, align: AlignmentType.CENTER, widthTwips: sec7ValCols[1], marginTB: 18 }),
                simpleCell(r.origBranchFsv, { size: SZ_SMALL, align: AlignmentType.CENTER, widthTwips: sec7ValCols[2], marginTB: 18 }),
                simpleCell(r.origSurvMv, { size: SZ_SMALL, align: AlignmentType.CENTER, widthTwips: sec7ValCols[3], marginTB: 18 }),
                simpleCell(r.origSurvFsv, { size: SZ_SMALL, align: AlignmentType.CENTER, widthTwips: sec7ValCols[4], marginTB: 18 }),
                simpleCell(r.presBranchMv, { size: SZ_SMALL, align: AlignmentType.CENTER, widthTwips: sec7ValCols[5], marginTB: 18 }),
                simpleCell(r.presBranchFsv, { size: SZ_SMALL, align: AlignmentType.CENTER, widthTwips: sec7ValCols[6], marginTB: 18 }),
                simpleCell(r.presSurvMv, { size: SZ_SMALL, align: AlignmentType.CENTER, widthTwips: sec7ValCols[7], marginTB: 18 }),
                simpleCell(r.presSurvFsv, { size: SZ_SMALL, align: AlignmentType.CENTER, widthTwips: sec7ValCols[8], marginTB: 18 }),
                simpleCell(r.presAvMv, { size: SZ_SMALL, align: AlignmentType.CENTER, widthTwips: sec7ValCols[9], marginTB: 18 }),
                simpleCell(r.presAvFsv, { size: SZ_SMALL, align: AlignmentType.CENTER, widthTwips: sec7ValCols[10], marginTB: 18 }),
              ],
            })
        ),
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell(
              `Outstanding to Lowest Security value Coverage (%): ${bp?.section7CoveragePct || '……………..%'}`,
              {
                bold: true,
                size: SZ_SMALL,
                columnSpan: 11,
                align: AlignmentType.CENTER,
                widthTwips: BODY_WIDTH_TWIPS,
                marginTB: 20,
              }
            ),
          ],
        }),
      ],
    }),
    p([txt('c. Others:', { bold: true })], { indentLeft: 120, before: 35, after: 10 }),
    p([txt('a. Usual Charge documents.', { bold: true })], { indentLeft: 320, before: 10, after: 10 }),
    p([txt('b. Acceptance of Branch Sanction by the Client.', { bold: true })], { indentLeft: 320, before: 10, after: 15 }),
    ...(memo.sec7OthersNote
      ? memo.sec7OthersNote
          .split('\n')
          .filter((l) => l.trim())
          .map((line) => p([txt(line.trim(), { size: SZ_BODY })], { indentLeft: 320, before: 10, after: 10 }))
      : [])
  );

  // Section 8 (Source of Repayment)
  body.push(
    p(
      [
        txt('8. Source of Repayment: ', { bold: true, underline: true }),
        txt(memo.sourceOfRepayment || ''),
      ],
      { before: 40, after: 30 }
    )
  );

  // Section 9 (Sequence of Head Office Sanction(S))
  const sec9Rows =
    bp?.section9Rows && bp.section9Rows.length > 0
      ? bp.section9Rows
      : ['a.', 'b.', 'c.', 'd.', 'e.', 'f.', 'g.'].map((sl) => ({
          sl,
          referenceNo: '',
          date: '',
          limit: '',
          remarks: '',
        }));
  const sec9Cols = pctWidths([7, 36, 16, 15, 26]);

  body.push(
    p([txt('9. Sequence of Head Office Sanction(S):', { bold: true, underline: true })], { before: 40, after: 20 }),
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: sec9Cols,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            simpleCell('SL', { bold: true, align: AlignmentType.CENTER, widthTwips: sec9Cols[0], marginTB: 22 }),
            simpleCell('HO Sanction Reference No.', { bold: true, widthTwips: sec9Cols[1], marginTB: 22 }),
            simpleCell('Date', { bold: true, align: AlignmentType.CENTER, widthTwips: sec9Cols[2], marginTB: 22 }),
            simpleCell('Limit', { bold: true, align: AlignmentType.CENTER, widthTwips: sec9Cols[3], marginTB: 22 }),
            simpleCell('Remarks', { bold: true, widthTwips: sec9Cols[4], marginTB: 22 }),
          ],
        }),
        ...sec9Rows.map(
          (r) =>
            new TableRow({
              cantSplit: true,
              children: [
                simpleCell(r.sl, { align: AlignmentType.CENTER, widthTwips: sec9Cols[0], marginTB: 20 }),
                simpleCell(r.referenceNo, { widthTwips: sec9Cols[1], marginTB: 20 }),
                simpleCell(r.date, { align: AlignmentType.CENTER, widthTwips: sec9Cols[2], marginTB: 20 }),
                simpleCell(r.limit, { align: AlignmentType.CENTER, widthTwips: sec9Cols[3], marginTB: 20 }),
                simpleCell(r.remarks, { widthTwips: sec9Cols[4], marginTB: 20 }),
              ],
            })
        ),
      ],
    })
  );

  // Section 10 (Brief Particulars of The Client) — matching rowSpan=2 on g. and h. and superscript 1st
  const sp = bp?.section10Particulars;
  const sec10Cols = pctWidths([6, 37, 4, 53]);
  const sec10Rows: TableRow[] = [
    new TableRow({
      cantSplit: true,
      children: [
        simpleCell('a.', { widthTwips: sec10Cols[0], marginTB: 16 }),
        simpleCell('Name of the Client', { widthTwips: sec10Cols[1], marginTB: 16 }),
        simpleCell(':', { widthTwips: sec10Cols[2], align: AlignmentType.CENTER, marginTB: 16 }),
        simpleCell(sp?.nameOfClient || '', { widthTwips: sec10Cols[3], marginTB: 16 }),
      ],
    }),
    new TableRow({
      cantSplit: true,
      children: [
        simpleCell('b.', { widthTwips: sec10Cols[0], marginTB: 16 }),
        simpleCell('Unique Customer ID', { widthTwips: sec10Cols[1], marginTB: 16 }),
        simpleCell(':', { widthTwips: sec10Cols[2], align: AlignmentType.CENTER, marginTB: 16 }),
        simpleCell(sp?.uniqueCustomerId || '', { widthTwips: sec10Cols[3], marginTB: 16 }),
      ],
    }),
    new TableRow({
      cantSplit: true,
      children: [
        simpleCell('c.', { widthTwips: sec10Cols[0], marginTB: 16 }),
        simpleCell('AWCD/MSD Account with opening date', { widthTwips: sec10Cols[1], marginTB: 16 }),
        simpleCell(':', { widthTwips: sec10Cols[2], align: AlignmentType.CENTER, marginTB: 16 }),
        simpleCell(sp?.awcdMsdAccount || '', { widthTwips: sec10Cols[3], marginTB: 16 }),
      ],
    }),
    new TableRow({
      cantSplit: true,
      children: [
        simpleCell('d.', { widthTwips: sec10Cols[0], marginTB: 16 }),
        simpleCell('Business Type/Nature', { widthTwips: sec10Cols[1], marginTB: 16 }),
        simpleCell(':', { widthTwips: sec10Cols[2], align: AlignmentType.CENTER, marginTB: 16 }),
        simpleCell(sp?.businessTypeNature || '', { widthTwips: sec10Cols[3], marginTB: 16 }),
      ],
    }),
    new TableRow({
      cantSplit: true,
      children: [
        simpleCell('e.', { widthTwips: sec10Cols[0], marginTB: 16 }),
        simpleCell('Business Management Type', { widthTwips: sec10Cols[1], marginTB: 16 }),
        simpleCell(':', { widthTwips: sec10Cols[2], align: AlignmentType.CENTER, marginTB: 16 }),
        simpleCell(sp?.businessManagementType || '', { widthTwips: sec10Cols[3], marginTB: 16 }),
      ],
    }),
    new TableRow({
      cantSplit: true,
      children: [
        simpleCell('f.', { widthTwips: sec10Cols[0], marginTB: 16 }),
        simpleCell('Client Type', { widthTwips: sec10Cols[1], marginTB: 16 }),
        simpleCell(':', { widthTwips: sec10Cols[2], align: AlignmentType.CENTER, marginTB: 16 }),
        simpleCell(sp?.clientType || '', { widthTwips: sec10Cols[3], marginTB: 16 }),
      ],
    }),
    // g. spans 2 rows
    new TableRow({
      cantSplit: true,
      children: [
        simpleCell('g.', { rowSpan: 2, widthTwips: sec10Cols[0], vAlign: VerticalAlign.CENTER, marginTB: 16 }),
        simpleCell('Key Person with Phone No.', { widthTwips: sec10Cols[1], marginTB: 16 }),
        simpleCell(':', { widthTwips: sec10Cols[2], align: AlignmentType.CENTER, marginTB: 16 }),
        simpleCell(sp?.keyPersonPhone || '', { widthTwips: sec10Cols[3], marginTB: 16 }),
      ],
    }),
    new TableRow({
      cantSplit: true,
      children: [
        simpleCell('Address', { widthTwips: sec10Cols[1], marginTB: 16 }),
        simpleCell(':', { widthTwips: sec10Cols[2], align: AlignmentType.CENTER, marginTB: 16 }),
        simpleCell(sp?.address || '', { widthTwips: sec10Cols[3], marginTB: 16 }),
      ],
    }),
    // h. spans 2 rows
    new TableRow({
      cantSplit: true,
      children: [
        simpleCell('h.', { rowSpan: 2, widthTwips: sec10Cols[0], vAlign: VerticalAlign.CENTER, marginTB: 16 }),
        simpleCell('Office/Business Address', { widthTwips: sec10Cols[1], marginTB: 16 }),
        simpleCell(':', { widthTwips: sec10Cols[2], align: AlignmentType.CENTER, marginTB: 16 }),
        simpleCell(sp?.officeBusinessAddress || '', { widthTwips: sec10Cols[3], marginTB: 16 }),
      ],
    }),
    new TableRow({
      cantSplit: true,
      children: [
        simpleCell('Factory Address', { widthTwips: sec10Cols[1], marginTB: 16 }),
        simpleCell(':', { widthTwips: sec10Cols[2], align: AlignmentType.CENTER, marginTB: 16 }),
        simpleCell(sp?.factoryAddress || '', { widthTwips: sec10Cols[3], marginTB: 16 }),
      ],
    }),
    ...[
      ['i.', 'Purpose of Investment', sp?.purposeOfInvestment || ''],
      ['j.', 'Limit of Investment', sp?.limitOfInvestment || ''],
      ['k.', 'Mode of Investment', sp?.modeOfInvestment || ''],
      ['l.', 'Last Visit Date', sp?.lastVisitDate || ''],
      ['m.', 'Stock Position', sp?.stockPosition || ''],
      ['n.', 'Fixed Asset', sp?.fixedAsset || ''],
      ['o.', 'Business Net Worth', sp?.businessNetWorth || ''],
      ['p.', 'Insurance Coverage', sp?.insuranceCoverage || ''],
      ['q.', 'Trade License & Validity', sp?.tradeLicenseValidity || ''],
      ['r', 'Other Regulatory Documents', sp?.otherRegulatoryDocuments || ''],
      ['s', 'Date of Original Investment', sp?.dateOfOriginalInvestment || ''],
    ].map(
      ([idx, label, val]) =>
        new TableRow({
          cantSplit: true,
          children: [
            simpleCell(idx, { widthTwips: sec10Cols[0], marginTB: 16 }),
            simpleCell(label, { widthTwips: sec10Cols[1], marginTB: 16 }),
            simpleCell(':', { widthTwips: sec10Cols[2], align: AlignmentType.CENTER, marginTB: 16 }),
            simpleCell(val, { widthTwips: sec10Cols[3], marginTB: 16 }),
          ],
        })
    ),
    // t. with superscript st
    new TableRow({
      cantSplit: true,
      children: [
        simpleCell('t.', { widthTwips: sec10Cols[0], marginTB: 16 }),
        cell(
          [
            p(
              [
                txt('1', { size: SZ_BODY }),
                txt('st', { size: SZ_BODY, superScript: true }),
                txt(' Disbursement Manager', { size: SZ_BODY }),
              ],
              { before: 10, after: 10, line: 225 }
            ),
          ],
          { widthTwips: sec10Cols[1], marginTB: 16 }
        ),
        simpleCell(':', { widthTwips: sec10Cols[2], align: AlignmentType.CENTER, marginTB: 16 }),
        simpleCell(sp?.firstDisbursementManager || '', { widthTwips: sec10Cols[3], marginTB: 16 }),
      ],
    }),
    new TableRow({
      cantSplit: true,
      children: [
        simpleCell('u.', { widthTwips: sec10Cols[0], marginTB: 16 }),
        simpleCell('First Date of Classification with Status', { widthTwips: sec10Cols[1], marginTB: 16 }),
        simpleCell(':', { widthTwips: sec10Cols[2], align: AlignmentType.CENTER, marginTB: 16 }),
        simpleCell(sp?.firstDateOfClassification || '', { widthTwips: sec10Cols[3], marginTB: 16 }),
      ],
    }),
  ];

  body.push(
    p([txt('10. Brief Particulars of The Client:', { bold: true, underline: true })], { before: 50, after: 20 }),
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: sec10Cols,
      rows: sec10Rows,
    })
  );

  // Particulars of Proprietor/Partner/Directors (2 default rows matching IC_Memo.pdf)
  const dirRows =
    bp?.section10Directors && bp.section10Directors.length > 0
      ? bp.section10Directors
      : [
          { name: '', age: '', education: '', designation: '', sharePct: '', experience: '', netWorth: '' },
          { name: '', age: '', education: '', designation: '', sharePct: '', experience: '', netWorth: '' },
        ];
  const dirCols = pctWidths([22, 8, 15, 16, 12, 15, 12]);
  body.push(
    p([txt('•  Particular of the Proprietor/Partner/Directors:', { bold: true, underline: true })], { before: 45, after: 20 }),
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: dirCols,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            simpleCell('Name of the Directors', { bold: true, widthTwips: dirCols[0], marginTB: 22 }),
            simpleCell('Age', { bold: true, align: AlignmentType.CENTER, widthTwips: dirCols[1], marginTB: 22 }),
            simpleCell('Education', { bold: true, widthTwips: dirCols[2], marginTB: 22 }),
            simpleCell('Designation', { bold: true, widthTwips: dirCols[3], marginTB: 22 }),
            simpleCell('Share (%)', { bold: true, align: AlignmentType.CENTER, widthTwips: dirCols[4], marginTB: 22 }),
            simpleCell('Experience', { bold: true, align: AlignmentType.CENTER, widthTwips: dirCols[5], marginTB: 22 }),
            simpleCell('Net worth', { bold: true, align: AlignmentType.CENTER, widthTwips: dirCols[6], marginTB: 22 }),
          ],
        }),
        ...dirRows.map(
          (d) =>
            new TableRow({
              cantSplit: true,
              children: [
                simpleCell(d.name, { widthTwips: dirCols[0], marginTB: 32 }),
                simpleCell(d.age, { align: AlignmentType.CENTER, widthTwips: dirCols[1], marginTB: 32 }),
                simpleCell(d.education, { widthTwips: dirCols[2], marginTB: 32 }),
                simpleCell(d.designation, { widthTwips: dirCols[3], marginTB: 32 }),
                simpleCell(d.sharePct, { align: AlignmentType.CENTER, widthTwips: dirCols[4], marginTB: 32 }),
                simpleCell(d.experience, { align: AlignmentType.CENTER, widthTwips: dirCols[5], marginTB: 32 }),
                simpleCell(d.netWorth, { align: AlignmentType.CENTER, widthTwips: dirCols[6], marginTB: 32 }),
              ],
            })
        ),
      ],
    })
  );

  // Remarks: Director of any Bank/FI/Insurance (3 default rows matching IC_Memo.pdf Page 4)
  const bankDirRows =
    bp?.section10BankDirectors && bp.section10BankDirectors.length > 0
      ? bp.section10BankDirectors
      : [
          { name: '', bankNbfi: '', designation: '', sharePct: '', numShares: '' },
          { name: '', bankNbfi: '', designation: '', sharePct: '', numShares: '' },
          { name: '', bankNbfi: '', designation: '', sharePct: '', numShares: '' },
        ];
  const bankDirCols = pctWidths([22, 22, 18, 19, 19]);
  body.push(
    p([txt('•  Remarks: Director of any Bank/ FI/ Insurance company:', { bold: true, underline: true })], {
      before: 45,
      after: 20,
    }),
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: bankDirCols,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            simpleCell('Name of the Director', { bold: true, widthTwips: bankDirCols[0], marginTB: 22 }),
            simpleCell('Name of the bank/NBFIs', { bold: true, widthTwips: bankDirCols[1], marginTB: 22 }),
            simpleCell('Designation/Position', { bold: true, widthTwips: bankDirCols[2], marginTB: 22 }),
            simpleCell('Percentage of Shares', { bold: true, align: AlignmentType.CENTER, widthTwips: bankDirCols[3], marginTB: 22 }),
            simpleCell('Number of Shares', { bold: true, align: AlignmentType.CENTER, widthTwips: bankDirCols[4], marginTB: 22 }),
          ],
        }),
        ...bankDirRows.map(
          (d) =>
            new TableRow({
              cantSplit: true,
              children: [
                simpleCell(d.name, { widthTwips: bankDirCols[0], marginTB: 32 }),
                simpleCell(d.bankNbfi, { widthTwips: bankDirCols[1], marginTB: 32 }),
                simpleCell(d.designation, { widthTwips: bankDirCols[2], marginTB: 32 }),
                simpleCell(d.sharePct, { align: AlignmentType.CENTER, widthTwips: bankDirCols[3], marginTB: 32 }),
                simpleCell(d.numShares, { align: AlignmentType.CENTER, widthTwips: bankDirCols[4], marginTB: 32 }),
              ],
            })
        ),
      ],
    })
  );

  // Section 11 (Financials of The Client — 2 default rows matching IC_Memo.pdf Page 4)
  const finRows =
    bp?.section11Rows && bp.section11Rows.length > 0
      ? bp.section11Rows
      : [
          {
            year: '',
            grossSales: bp?.section11FinancialsNote || '',
            grossProfit: '',
            netProfit: '',
            totalAsset: '',
            totalLiabilities: '',
            netWorth: '',
          },
          {
            year: '',
            grossSales: '',
            grossProfit: '',
            netProfit: '',
            totalAsset: '',
            totalLiabilities: '',
            netWorth: '',
          },
        ];
  const finCols = pctWidths([10, 15, 15, 15, 15, 15, 15]);
  body.push(
    p([txt('11. Financials of The Client:', { bold: true, underline: true })], { before: 50, after: 20 }),
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: finCols,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            simpleCell('Year', { bold: true, align: AlignmentType.CENTER, widthTwips: finCols[0], marginTB: 22 }),
            simpleCell('Gross Sales', { bold: true, align: AlignmentType.CENTER, widthTwips: finCols[1], marginTB: 22 }),
            simpleCell('Gross Profit', { bold: true, align: AlignmentType.CENTER, widthTwips: finCols[2], marginTB: 22 }),
            simpleCell('Net Profit', { bold: true, align: AlignmentType.CENTER, widthTwips: finCols[3], marginTB: 22 }),
            simpleCell('Total Asset', { bold: true, align: AlignmentType.CENTER, widthTwips: finCols[4], marginTB: 22 }),
            simpleCell('Total liabilities', { bold: true, align: AlignmentType.CENTER, widthTwips: finCols[5], marginTB: 22 }),
            simpleCell('Net Worth', { bold: true, align: AlignmentType.CENTER, widthTwips: finCols[6], marginTB: 22 }),
          ],
        }),
        ...finRows.map(
          (r) =>
            new TableRow({
              cantSplit: true,
              children: [
                simpleCell(r.year, { align: AlignmentType.CENTER, widthTwips: finCols[0], marginTB: 32 }),
                simpleCell(r.grossSales, { align: AlignmentType.CENTER, widthTwips: finCols[1], marginTB: 32 }),
                simpleCell(r.grossProfit, { align: AlignmentType.CENTER, widthTwips: finCols[2], marginTB: 32 }),
                simpleCell(r.netProfit, { align: AlignmentType.CENTER, widthTwips: finCols[3], marginTB: 32 }),
                simpleCell(r.totalAsset, { align: AlignmentType.CENTER, widthTwips: finCols[4], marginTB: 32 }),
                simpleCell(r.totalLiabilities, { align: AlignmentType.CENTER, widthTwips: finCols[5], marginTB: 32 }),
                simpleCell(r.netWorth, { align: AlignmentType.CENTER, widthTwips: finCols[6], marginTB: 32 }),
              ],
            })
        ),
      ],
    })
  );

  // Section 12 (Contents of the Client's Prayer)
  body.push(
    p(
      [
        txt('12. Contents of the Client’s Prayer: ', { bold: true, underline: true }),
        txt(memo.clientPrayer || ''),
      ],
      { before: 50, after: 25 }
    )
  );

  // Section 12 (Branch Visit Report)
  body.push(
    p(
      [
        txt('12. Branch Visit Report: ', { bold: true, underline: true }),
        txt(memo.branchVisitReport || ''),
      ],
      { before: 35, after: 25 }
    )
  );

  // Section 13 (Recommendation of The Branch)
  body.push(
    p(
      [
        txt('13. Recommendation of The Branch: ', { bold: true, underline: true }),
        txt(`The Branch Recommended the following vide Proposal Ref: ${v('#64#')} dated ${v('#65#')}`),
      ],
      { before: 35, after: 20 }
    )
  );
  if (memo.sec13BranchRecommendation) {
    for (const line of memo.sec13BranchRecommendation.split('\n').filter((l) => l.trim())) {
      body.push(p([txt(line.trim())], { indentLeft: 180, before: 12, after: 15 }));
    }
  }

  // Section 14
  body.push(
    p(
      [
        txt('14. Recommendation of Internal Control & Compliance Division: ', { bold: true, underline: true }),
        memo.sec14IccdRecommendation
          ? txt(memo.sec14IccdRecommendation)
          : txt('N/A', { bold: true }),
      ],
      { before: 35, after: 25 }
    )
  );

  // Section 15 (Observation and Recommendation of SAMD)
  body.push(
    p([txt('15. Observation and Recommendation of SAMD:', { bold: true, underline: true })], { before: 35, after: 15 })
  );
  if (memo.samdObservation) {
    for (const line of memo.samdObservation.split('\n').filter((l) => l.trim())) {
      body.push(p([txt(line.trim())], { indentLeft: 180, before: 15, after: 15 }));
    }
  }

  for (const sec of insertedAfterSec15) {
    body.push(
      p([txt(`${sec.sectionNumber}. ${sec.heading}:`, { bold: true, underline: true })], {
        before: 45,
        after: 15,
      })
    );
    if (sec.text && sec.text.trim()) {
      for (const line of sec.text.split('\n').filter((l) => l.trim())) {
        body.push(p([txt(line.trim())], { indentLeft: 180, before: 12, after: 16, line: 245 }));
      }
    }
  }

  // Particulars of the facility (Section 16 in IC Memo; adjusted cleanly in MEC/Board Memos)
  const sec16Cols = pctWidths([5, 24, 71]);
  const sec16Rows: Array<[string, string, string, boolean]> = [
    ['a.', 'Facility', 'EXIT with Waiver under BRPD 1 Circular Letter No. 23 dated 29.06.2026', true],
    ['b.', 'Outstanding', `Tk. ${v('#25#')} million as on ${v('#68#')}.`, true],
    ['c.', 'Down Payment', `Down-payment of Tk. ${v('#35#')} Million is kept in ${v('#36#')} A/C ${v('#36#')}`, true],
    ['d.', 'Waiver', `Tk. ${v('#4#')} million as on ${v('#68#')}`, true],
    ['e.', 'Settlement/Exit amount', `Tk. ${v('#27#')} Million as on ${v('#68#')}`, true],
    ['f.', 'Settlement period', `${v('#28#')}`, true],
    ['g.', 'Profit/Rent rate', 'Without charging further profit.', false],
    ['h.', 'Repayment Method', 'Lump sum.', false],
  ];
  body.push(
    p([txt(`${particularsSecNum}. Particulars of the facility:`, { bold: true, underline: true })], { before: 45, after: 15 }),
    new Table({
      width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
      columnWidths: sec16Cols,
      rows: sec16Rows.map(
        ([sl, label, val, isBold]) =>
          new TableRow({
            cantSplit: true,
            children: [
              simpleCell(sl, { align: AlignmentType.CENTER, widthTwips: sec16Cols[0], marginTB: 20 }),
              simpleCell(label, { widthTwips: sec16Cols[1], marginTB: 20 }),
              simpleCell(val, { bold: isBold, widthTwips: sec16Cols[2], marginTB: 20 }),
            ],
          })
      ),
    })
  );

  // Section 17 (Special Term(s) & Condition(s))
  const terms: Array<[string, string, boolean]> = [
    [
      'i.',
      'The waiver of profit shall be adjusted from FAD, Head Office from the provision amount kept (if any) and the rest amount shall be adjusted by branch gradually after adjustment of the settlement amount within the approved time frame.',
      false,
    ],
    [
      'ii.',
      'The final settlement amount will be determined at the date of execution of settlement considering up to date liability position.',
      false,
    ],
    [
      'iii.',
      'The Branch must ensure adjustment of all other direct and indirect liabilities of the client before allowing the final redemption of the mortgaged collateral security(s).',
      false,
    ],
    [
      'iv.',
      'The Branch must ensure adjustment of all non-funded liabilities (L/C, Bank Guarantee, Credit Card etc.) before allowing final redemption of the mortgaged collateral security(s). The Branch must also ensure that any encashable security (Margin for Bank Guarantee, L/C, MTDR, Scheme deposit in the name of the client) including collateral is properly mentioned in the proposal. If any encashable security (Margin for Bank Guarantee, L/C, MTDR, Scheme deposit in the name of the client) is found not mentioned in the proposal and later encashed in favor of the client, the Manager will be held responsible.',
      false,
    ],
    [
      'v.',
      'If the client fails to repay or fails to adjust the liability within the given time or fails to pay installments regularly, all facilities shall be cancelled and legal action shall be initiated by the Branch.',
      false,
    ],
    [
      'vi.',
      'Branch should be utmost sincere to protect bank interest in the course of execution partial redemption.',
      false,
    ],
    [
      'vii.',
      'Branch is advised to follow the guidelines of Bangladesh Bank BRPD circular No. 06/22, Circular Letter No.18/22, 46/22 and subsequent up-dates as well as all instruction circular(s) of SIBL.',
      false,
    ],
    [
      'viii.',
      'The Profit rate of the stipulated investment shall not be changed or reduced to 0.00% until full and final adjustment of the Investment, as approved by Head Office.',
      true,
    ],
    [
      'ix.',
      'All the legal expenses and other charges (if any) already incurred and to be incurred in connection with the withdrawal of case/suits and others issues shall be borne by the client.',
      false,
    ],
  ];

  body.push(
    p([txt(`${termsSecNum}. Special Term(s) & Condition(s):`, { bold: true, underline: true })], { before: 50, after: 15 }),
    p(
      [txt('If the proposal approved, the branch is advised to comply the following special terms and conditions:')],
      { indentLeft: 280, before: 15, after: 20 }
    ),
    ...terms.map(([num, clause, isBold]) =>
      p([txt(`${num}  `, { bold: isBold }), txt(clause, { bold: isBold })], {
        indentLeft: 520,
        before: 25,
        after: 25,
        line: 260,
      })
    )
  );

  if (options.includeClosingBlock !== false) {
    // Closing box and Signatories
    const sigCols = pctWidths([50, 50]);
    body.push(
      p([], { before: 50, after: 30 }),
      new Table({
        width: { size: BODY_WIDTH_TWIPS, type: WidthType.DXA },
        layout: TableLayoutType.FIXED,
        columnWidths: [BODY_WIDTH_TWIPS],
        rows: [
          new TableRow({
            cantSplit: true,
            children: [
              simpleCell('The proposal is hereby placed before the Investment Committee for kind recommendation.', {
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
      p([txt('Special Asset Management Division', { bold: true })], { before: 0, after: 50 }),
      p([txt('Submitted for kind decision.', { bold: true })], { align: AlignmentType.CENTER, before: 40, after: 100 }),
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
                  p([txt('(Shahidur Rahman khan)', { bold: true })], { before: 200, after: 10 }),
                  p([txt('Deputy Managing Director & CRO')], { before: 0, after: 10 }),
                  p([txt('Sammilito Islami Bank PLC.')], { before: 0, after: 10 }),
                ],
                { widthTwips: sigCols[0], borders: ALL_NO_BORDERS }
              ),
              cell(
                [
                  p([txt('(Ahmed Zafrul Hasan)', { bold: true })], { before: 200, after: 10 }),
                  p([txt('Additional Managing Director & COO')], { before: 0, after: 10 }),
                  p([txt('Sammilito Islami Bank PLC.')], { before: 0, after: 10 }),
                ],
                { widthTwips: sigCols[1], borders: ALL_NO_BORDERS }
              ),
            ],
          }),
        ],
      }),
      p([txt('(Abedur Rahman Sikder)', { bold: true })], { align: AlignmentType.CENTER, before: 240, after: 10 }),
      p([txt('Managing Director')], { align: AlignmentType.CENTER, before: 0, after: 10 }),
      p([txt('Sammilito Islami Bank PLC.')], { align: AlignmentType.CENTER, before: 0, after: 10 }),
      p([txt('(Transferee of the assets and liabilities of')], { align: AlignmentType.CENTER, before: 0, after: 10 }),
      p([txt('EXIM Bank, SIBL, FSIB, Global Islami')], { align: AlignmentType.CENTER, before: 0, after: 10 }),
      p([txt('Bank and Union Bank)')], { align: AlignmentType.CENTER, before: 0, after: 10 })
    );
  }

  return body;
}
