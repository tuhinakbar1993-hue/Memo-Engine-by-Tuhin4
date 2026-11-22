import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import { cleanText } from '../utils/numberFormat';

export interface StructuredTable {
  rows: string[][];
}

export interface ExtractedDocxContent {
  paragraphs: string[];
  tables: StructuredTable[];
  lines: string[];
  rawText: string;
}

export interface ExcelSheetContent {
  name: string;
  /** Formatted text values for every cell (dates formatted DD/MM/YYYY, formula results resolved) */
  rows: string[][];
  /** Raw typed values (numbers preserved as exact floats, dates formatted DD/MM/YYYY, strings trimmed) */
  typedRows: Array<Array<string | number | null>>;
}

export interface ExtractedExcelContent {
  sheets: ExcelSheetContent[];
  lines: string[];
  rawText: string;
}

const docxCache = new Map<string, ExtractedDocxContent>();
const excelCache = new Map<string, ExtractedExcelContent>();

function fileSig(file: File): string {
  return `${file.name}:${file.size}:${file.lastModified}`;
}

export function isDocxFile(file: File): boolean {
  const name = file.name.toLowerCase();
  return (
    name.endsWith('.docx') ||
    file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  );
}

export function isExcelFile(file: File): boolean {
  const name = file.name.toLowerCase();
  return (
    name.endsWith('.xlsx') ||
    name.endsWith('.xls') ||
    name.endsWith('.csv') ||
    file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
    file.type === 'application/vnd.ms-excel' ||
    file.type === 'text/csv'
  );
}

export function isPdfFile(file: File): boolean {
  const name = file.name.toLowerCase();
  return name.endsWith('.pdf') || file.type === 'application/pdf';
}

function extractParagraphText(pEl: Element): string {
  const parts: string[] = [];
  const walk = (node: Node) => {
    if (node.nodeType === 1) {
      const el = node as Element;
      const local = el.localName;
      if (local === 't' || local === 'delText' || local === 'instrText') {
        if (local === 't' && el.textContent) {
          parts.push(el.textContent);
        }
        return;
      }
      if (local === 'br' || local === 'cr') {
        parts.push('\n');
        return;
      }
      if (local === 'tab') {
        parts.push(' ');
        return;
      }
    }
    for (let i = 0; i < node.childNodes.length; i++) {
      walk(node.childNodes[i]);
    }
  };
  walk(pEl);
  return parts
    .join('')
    .replace(/[ \t]+/g, ' ')
    .replace(/\s*\n\s*/g, '\n')
    .trim();
}

function directChildrenByLocalName(parent: Element, name: string): Element[] {
  const res: Element[] = [];
  for (let i = 0; i < parent.childNodes.length; i++) {
    const n = parent.childNodes[i];
    if (n.nodeType === 1 && (n as Element).localName === name) {
      res.push(n as Element);
    }
  }
  return res;
}

function firstDescendantByLocalName(parent: Element, name: string): Element | null {
  for (let i = 0; i < parent.childNodes.length; i++) {
    const n = parent.childNodes[i];
    if (n.nodeType === 1) {
      const el = n as Element;
      if (el.localName === name) return el;
      const found = firstDescendantByLocalName(el, name);
      if (found) return found;
    }
  }
  return null;
}

function extractCellText(tcEl: Element): string {
  const paragraphs: string[] = [];
  for (let i = 0; i < tcEl.childNodes.length; i++) {
    const n = tcEl.childNodes[i];
    if (n.nodeType === 1) {
      const el = n as Element;
      if (el.localName === 'p') {
        const t = extractParagraphText(el);
        if (t) paragraphs.push(t);
      } else if (el.localName === 'tbl') {
        const nested = extractTable(el);
        for (const r of nested.rows) {
          const line = r.filter(Boolean).join(' | ');
          if (line) paragraphs.push(line);
        }
      }
    }
  }
  return paragraphs.join('\n').trim();
}

function extractTable(tblEl: Element): StructuredTable {
  const trEls = directChildrenByLocalName(tblEl, 'tr');
  const rows: string[][] = [];
  const activeVMerge: Array<string | null> = [];

  for (const tr of trEls) {
    const tcEls = directChildrenByLocalName(tr, 'tc');
    const rowCells: string[] = [];
    let colIdx = 0;

    for (const tc of tcEls) {
      // Advance past any column currently vertically merged from above
      while (activeVMerge[colIdx] !== undefined && activeVMerge[colIdx] !== null) {
        // Check if this cell corresponds to continuation; actually Word emits explicit <w:tc> with empty vMerge for continued cells!
        break;
      }

      const tcPr = firstDescendantByLocalName(tc, 'tcPr');
      let span = 1;
      let vMergeVal: string | null = null;
      let hasVMerge = false;

      if (tcPr) {
        const gridSpan = firstDescendantByLocalName(tcPr, 'gridSpan');
        if (gridSpan) {
          const valAttr = gridSpan.getAttribute('w:val') ?? gridSpan.getAttribute('val');
          const parsedSpan = Number(valAttr);
          if (Number.isFinite(parsedSpan) && parsedSpan > 1) {
            span = parsedSpan;
          }
        }
        const vMerge = firstDescendantByLocalName(tcPr, 'vMerge');
        if (vMerge) {
          hasVMerge = true;
          vMergeVal = vMerge.getAttribute('w:val') ?? vMerge.getAttribute('val') ?? 'continue';
        }
      }

      let cellText = extractCellText(tc);
      if (hasVMerge && vMergeVal !== 'restart' && !cellText && activeVMerge[colIdx]) {
        // Carry forward header label only for continuation cells when helpful? Keep raw cell empty for data rows,
        // so downstream parsers don't double-count numbers!
        cellText = '';
      }

      if (hasVMerge && vMergeVal === 'restart') {
        activeVMerge[colIdx] = cellText;
      } else if (!hasVMerge) {
        activeVMerge[colIdx] = null;
      }

      rowCells.push(cellText);
      for (let s = 1; s < span; s++) {
        rowCells.push('');
      }
      colIdx += span;
    }

    if (rowCells.some((c) => c !== '')) {
      rows.push(rowCells);
    } else {
      rows.push(rowCells);
    }
  }

  return { rows };
}

/**
 * Fast, cached reader for Microsoft Word (.docx) files.
 * Extracts paragraphs, headers/footers, and structured tables in reading order.
 */
export async function extractDocxContent(file: File): Promise<ExtractedDocxContent> {
  const sig = fileSig(file);
  const cached = docxCache.get(sig);
  if (cached) return cached;

  const buffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(buffer);

  const parser = new DOMParser();
  const paragraphs: string[] = [];
  const tables: StructuredTable[] = [];
  const orderedBlocks: string[] = [];

  // Parse header XML files if present (bank/branch names or print dates sometimes appear there)
  const headerPaths = Object.keys(zip.files)
    .filter((p) => /^word\/header\d*\.xml$/i.test(p))
    .sort();
  for (const hp of headerPaths) {
    const xmlStr = await zip.files[hp].async('string');
    const doc = parser.parseFromString(xmlStr, 'application/xml');
    const root = doc.documentElement;
    for (let i = 0; i < root.childNodes.length; i++) {
      const node = root.childNodes[i];
      if (node.nodeType !== 1) continue;
      const el = node as Element;
      if (el.localName === 'p') {
        const text = extractParagraphText(el);
        if (text) {
          paragraphs.push(text);
          orderedBlocks.push(text);
        }
      } else if (el.localName === 'tbl') {
        const tbl = extractTable(el);
        tables.push(tbl);
        for (const row of tbl.rows) {
          const line = row.map((c) => cleanText(c)).filter(Boolean).join(' | ');
          if (line) orderedBlocks.push(line);
        }
      }
    }
  }

  const docFile = zip.file('word/document.xml');
  if (!docFile) {
    throw new Error('Invalid .docx file: missing word/document.xml.');
  }

  const docXmlStr = await docFile.async('string');
  const xmlDoc = parser.parseFromString(docXmlStr, 'application/xml');
  const bodyEl = firstDescendantByLocalName(xmlDoc.documentElement, 'body') ?? xmlDoc.documentElement;

  for (let i = 0; i < bodyEl.childNodes.length; i++) {
    const node = bodyEl.childNodes[i];
    if (node.nodeType !== 1) continue;
    const el = node as Element;
    if (el.localName === 'p') {
      const text = extractParagraphText(el);
      if (text) {
        paragraphs.push(text);
        orderedBlocks.push(text);
      }
    } else if (el.localName === 'tbl') {
      const tbl = extractTable(el);
      tables.push(tbl);
      orderedBlocks.push('[TABLE]');
      for (const row of tbl.rows) {
        const line = row.map((c) => cleanText(c)).join(' | ');
        if (line.replace(/\|/g, '').trim()) {
          orderedBlocks.push(line);
        }
      }
      orderedBlocks.push('[/TABLE]');
    }
  }

  const lines = orderedBlocks.filter((l) => l !== '[TABLE]' && l !== '[/TABLE]');
  const result: ExtractedDocxContent = {
    paragraphs,
    tables,
    lines,
    rawText: orderedBlocks.join('\n'),
  };

  docxCache.set(sig, result);
  return result;
}

function formatDateValue(d: Date): string {
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

/**
 * Fast, cached reader for Microsoft Excel (.xlsx, .xls, .csv) files.
 * Reads all worksheets, evaluates formula outputs, handles dates & merged cells cleanly.
 */
export async function extractExcelContent(file: File): Promise<ExtractedExcelContent> {
  const sig = fileSig(file);
  const cached = excelCache.get(sig);
  if (cached) return cached;

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, {
    type: 'array',
    cellDates: true,
    cellNF: true,
    cellFormula: true,
  });

  const sheets: ExcelSheetContent[] = [];
  const allLines: string[] = [];
  const blockLines: string[] = [];

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet || !sheet['!ref']) continue;

    const range = XLSX.utils.decode_range(sheet['!ref']);
    const rowCount = range.e.r - range.s.r + 1;
    const colCount = range.e.c - range.s.c + 1;

    const rows: string[][] = Array.from({ length: rowCount }, () => Array(colCount).fill(''));
    const typedRows: Array<Array<string | number | null>> = Array.from({ length: rowCount }, () =>
      Array(colCount).fill(null)
    );

    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = sheet[addr] as XLSX.CellObject | undefined;
        if (!cell) continue;

        const ri = r - range.s.r;
        const ci = c - range.s.c;

        if (cell.v instanceof Date && !Number.isNaN(cell.v.getTime())) {
          const rawW = cell.w?.trim() ?? '';
          const hasFourDigitYear = /\d{4}/.test(rawW);
          const formatted = rawW && !/^#+$/.test(rawW) && hasFourDigitYear ? rawW : formatDateValue(cell.v);
          rows[ri][ci] = formatted;
          typedRows[ri][ci] = formatted;
        } else if (cell.t === 'n' && typeof cell.v === 'number') {
          // Check if numeric date serial using SheetJS built-in SSF date detector
          const isDateFmt = typeof cell.z === 'string' && XLSX.SSF.is_date(cell.z);
          if (isDateFmt && cell.w && !/^#+$/.test(cell.w.trim())) {
            rows[ri][ci] = cell.w.trim();
            typedRows[ri][ci] = cell.w.trim();
          } else {
            const disp = cell.w && !/^#+$/.test(cell.w.trim()) ? cell.w.trim() : String(cell.v);
            rows[ri][ci] = disp;
            typedRows[ri][ci] = cell.v;
          }
        } else if (cell.t === 'b') {
          const bStr = cell.v ? 'TRUE' : 'FALSE';
          rows[ri][ci] = bStr;
          typedRows[ri][ci] = bStr;
        } else {
          const s = cleanText(cell.w ?? (cell.v !== undefined && cell.v !== null ? String(cell.v) : ''));
          rows[ri][ci] = s;
          typedRows[ri][ci] = s || null;
        }
      }
    }

    // Drop completely empty trailing rows while preserving header/table alignment
    const nonEmptyRows: string[][] = [];
    const nonEmptyTyped: Array<Array<string | number | null>> = [];
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].some((v) => v !== '')) {
        nonEmptyRows.push(rows[i]);
        nonEmptyTyped.push(typedRows[i]);
      }
    }

    sheets.push({ name: sheetName, rows: nonEmptyRows, typedRows: nonEmptyTyped });

    blockLines.push(`[SHEET: ${sheetName}]`);
    for (const r of nonEmptyRows) {
      const joined = r.filter(Boolean).join(' | ');
      if (joined) {
        allLines.push(joined);
        blockLines.push(r.join(' | '));
      }
    }
  }

  const result: ExtractedExcelContent = {
    sheets,
    lines: allLines,
    rawText: blockLines.join('\n'),
  };

  excelCache.set(sig, result);
  return result;
}
