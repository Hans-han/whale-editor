import { loadPackage, parseXml, getNestedValue } from '../ooxmlPackage.js';
import type {
  DocxManifest,
  MeasurementInfo,
  ParagraphInfo,
  ParagraphFormatInfo,
  RunInfo,
  TableGeometryInfo,
  TableInfo,
  TableRowInfo,
  TableCellInfo,
  HeaderFooterInfo,
  CommentInfo,
  RelationshipSummary,
  MediaSummary,
  PackageValidation,
} from '../../types/index.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

interface XmlNode {
  [key: string]: unknown;
}

function toArray<T>(val: T | T[] | undefined): T[] {
  if (!val) return [];
  return Array.isArray(val) ? val : [val];
}

function extractText(node: unknown): string {
  if (!node) return '';
  if (typeof node === 'string') return node;
  if (typeof node !== 'object') return String(node);

  const obj = node as Record<string, unknown>;
  if ('w:t' in obj) {
    const t = obj['w:t'];
    if (typeof t === 'string') return t;
    if (t && typeof t === 'object' && '#text' in (t as Record<string, unknown>)) {
      return String((t as Record<string, unknown>)['#text']);
    }
  }

  let result = '';
  for (const key of Object.keys(obj)) {
    if (key.startsWith('@_')) continue;
    const val = obj[key];
    if (Array.isArray(val)) {
      for (const item of val) {
        result += extractText(item);
      }
    } else if (val && typeof val === 'object') {
      result += extractText(val);
    }
  }
  return result;
}

function getStyleId(para: XmlNode): string | undefined {
  const pPr = para['w:pPr'] as XmlNode | undefined;
  if (!pPr) return undefined;
  const pStyle = pPr['w:pStyle'] as XmlNode | undefined;
  return pStyle?.['@_w:val'] as string | undefined;
}

function getHeadingLevel(styleId: string | undefined, styles: Map<string, number>): number | undefined {
  if (!styleId) return undefined;
  return styles.get(styleId);
}

interface StyleMaps {
  headingLevels: Map<string, number>;
  names: Map<string, string>;
}

function readVal(node: unknown, attr = '@_w:val'): string | undefined {
  if (!node || typeof node !== 'object') return undefined;
  const value = (node as XmlNode)[attr];
  if (value === undefined || value === null) return undefined;
  return String(value);
}

function readNum(node: unknown, attr = '@_w:val'): number | undefined {
  const raw = readVal(node, attr);
  if (raw === undefined) return undefined;
  const n = Number(raw);
  return Number.isFinite(n) ? n : undefined;
}

function boolProp(node: unknown): boolean | undefined {
  if (node === '') return true;
  if (node === true) return true;
  if (node === false) return false;
  if (!node || typeof node !== 'object') return undefined;
  const value = (node as XmlNode)['@_w:val'];
  if (value === undefined) return true;
  if (value === false || value === 0) return false;
  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (normalized === 'false' || normalized === '0' || normalized === 'off' || normalized === 'no') {
      return false;
    }
  }
  return true;
}

function buildStyleMaps(stylesXml: string | null): StyleMaps {
  const headingLevels = new Map<string, number>();
  const names = new Map<string, string>();
  if (!stylesXml) return { headingLevels, names };

  const parsed = parseXml(stylesXml) as XmlNode;
  const styles = parsed['w:styles'] as XmlNode | undefined;
  if (!styles) return { headingLevels, names };

  const styleList = toArray(styles['w:style'] as XmlNode | XmlNode[]);
  for (const style of styleList) {
    const styleId = style['@_w:styleId'] as string | undefined;
    if (!styleId) continue;

    const name = style['w:name'] as XmlNode | undefined;
    const nameVal = name?.['@_w:val'] as string | undefined;
    if (nameVal) names.set(styleId, nameVal);
    if (nameVal?.startsWith('heading')) {
      const level = parseInt(nameVal.replace('heading', ''), 10);
      if (!isNaN(level)) {
        headingLevels.set(styleId, level);
      }
    }
  }
  return { headingLevels, names };
}

function generateId(prefix: string, index: number): string {
  return `${prefix}.${String(index).padStart(4, '0')}`;
}

function attrToString(value: unknown): string | undefined {
  if (value === undefined || value === null) return undefined;
  return String(value);
}

function extractCommentIds(para: XmlNode): string[] {
  const ids = new Set<string>();

  const walk = (node: unknown): void => {
    if (!node || typeof node !== 'object') return;
    if (Array.isArray(node)) {
      for (const item of node) walk(item);
      return;
    }

    const obj = node as XmlNode;
    const start = obj['w:commentRangeStart'] as XmlNode | undefined;
    const reference = obj['w:commentReference'] as XmlNode | undefined;
    const startId = attrToString(start?.['@_w:id']);
    const refId = attrToString(reference?.['@_w:id']);
    if (startId) ids.add(startId);
    if (refId) ids.add(refId);

    for (const key of Object.keys(obj)) {
      if (key.startsWith('@_')) continue;
      walk(obj[key]);
    }
  };

  walk(para);
  return [...ids];
}

function extractNumbering(para: XmlNode): { numId?: string; level?: string } | undefined {
  const pPr = para['w:pPr'] as XmlNode | undefined;
  const numPr = pPr?.['w:numPr'] as XmlNode | undefined;
  if (!numPr) return undefined;
  const numId = readVal(numPr['w:numId']);
  const level = readVal(numPr['w:ilvl']);
  if (!numId && !level) return undefined;
  return { numId, level };
}

function extractParagraphFormat(para: XmlNode): ParagraphFormatInfo | undefined {
  const pPr = para['w:pPr'] as XmlNode | undefined;
  if (!pPr) return undefined;

  const spacing = pPr['w:spacing'] as XmlNode | undefined;
  const ind = pPr['w:ind'] as XmlNode | undefined;
  const format: ParagraphFormatInfo = {
    alignment: readVal(pPr['w:jc']),
  };

  if (spacing) {
    format.spacing = {
      before: readNum(spacing, '@_w:before'),
      after: readNum(spacing, '@_w:after'),
      line: readNum(spacing, '@_w:line'),
      lineRule: readVal(spacing, '@_w:lineRule'),
    };
  }

  if (ind) {
    format.indentation = {
      left: readNum(ind, '@_w:left'),
      right: readNum(ind, '@_w:right'),
      firstLine: readNum(ind, '@_w:firstLine'),
      hanging: readNum(ind, '@_w:hanging'),
    };
  }

  const hasFormat =
    format.alignment ||
    Object.values(format.spacing ?? {}).some((v) => v !== undefined) ||
    Object.values(format.indentation ?? {}).some((v) => v !== undefined);

  return hasFormat ? format : undefined;
}

function extractRunText(run: XmlNode): string {
  const t = run['w:t'];
  if (typeof t === 'string') return t;
  if (Array.isArray(t)) {
    return t.map((item) => extractText({ 'w:t': item })).join('');
  }
  if (t && typeof t === 'object' && '#text' in (t as XmlNode)) {
    return String((t as XmlNode)['#text'] ?? '');
  }
  return extractText(run);
}

function extractRuns(para: XmlNode, paragraphId: string): RunInfo[] {
  const runs = toArray(para['w:r'] as XmlNode | XmlNode[]);
  return runs
    .map((run, index): RunInfo | null => {
      const text = extractRunText(run);
      const rPr = run['w:rPr'] as XmlNode | undefined;
      const fonts = rPr?.['w:rFonts'] as XmlNode | undefined;
      const underline = rPr?.['w:u'] as XmlNode | undefined;
      const color = rPr?.['w:color'] as XmlNode | undefined;
      const size = rPr?.['w:sz'] as XmlNode | undefined;
      const style = rPr?.['w:rStyle'] as XmlNode | undefined;
      const info: RunInfo = {
        runId: `${paragraphId}.r.${String(index + 1).padStart(4, '0')}`,
        textPreview: text.slice(0, 100),
        styleId: readVal(style),
        font: readVal(fonts, '@_w:ascii') ?? readVal(fonts, '@_w:hAnsi') ?? readVal(fonts, '@_w:eastAsia'),
        fontSizeHalfPoints: readNum(size),
        bold: boolProp(rPr?.['w:b']),
        italic: boolProp(rPr?.['w:i']),
        underline: underline ? readVal(underline) !== 'none' : undefined,
        color: readVal(color),
      };
      const hasTextOrStyle =
        info.textPreview ||
        info.styleId ||
        info.font ||
        info.fontSizeHalfPoints !== undefined ||
        info.bold !== undefined ||
        info.italic !== undefined ||
        info.underline !== undefined ||
        info.color;
      return hasTextOrStyle ? info : null;
    })
    .filter((r): r is RunInfo => r !== null);
}

function extractMeasurement(node: unknown): MeasurementInfo | undefined {
  if (!node || typeof node !== 'object') return undefined;
  const obj = node as XmlNode;
  const value = readNum(obj, '@_w:w');
  const type = readVal(obj, '@_w:type');
  if (value === undefined && !type) return undefined;
  return { type, value };
}

function extractTableGeometry(tbl: XmlNode): TableGeometryInfo | undefined {
  const tblPr = tbl['w:tblPr'] as XmlNode | undefined;
  const tblGrid = tbl['w:tblGrid'] as XmlNode | undefined;
  const gridCols = toArray(tblGrid?.['w:gridCol'] as XmlNode | XmlNode[])
    .map((col) => readNum(col, '@_w:w'))
    .filter((n): n is number => n !== undefined);

  const geometry: TableGeometryInfo = {
    width: extractMeasurement(tblPr?.['w:tblW']),
    gridColumns: gridCols.length > 0 ? gridCols : undefined,
  };

  return geometry.width || geometry.gridColumns ? geometry : undefined;
}

async function extractComments(
  pkg: Awaited<ReturnType<typeof loadPackage>>,
  anchors: Map<string, string[]>
): Promise<CommentInfo[]> {
  const commentsXml = await pkg.getContent('word/comments.xml');
  if (!commentsXml) return [];

  const parsed = parseXml(commentsXml) as XmlNode;
  const root = parsed['w:comments'] as XmlNode | undefined;
  if (!root) return [];

  const comments = toArray(root['w:comment'] as XmlNode | XmlNode[]);
  return comments
    .map((comment, index): CommentInfo | null => {
      const rawId = attrToString(comment['@_w:id']);
      const id = rawId ?? String(index);
      const text = extractText(comment).trim();
      if (!text) return null;
      return {
        commentId: `docx.comment.${id}`,
        author: attrToString(comment['@_w:author']),
        date: attrToString(comment['@_w:date']),
        textPreview: text.slice(0, 240),
        anchoredParagraphIds: anchors.get(id) ?? [],
      };
    })
    .filter((c): c is CommentInfo => c !== null);
}

export async function generateDocxManifest(buffer: Buffer): Promise<DocxManifest> {
  const pkg = await loadPackage(buffer);
  const paths = pkg.listPaths();

  const documentXml = await pkg.getContent('word/document.xml');
  if (!documentXml) {
    throw new Error('Invalid DOCX: missing word/document.xml');
  }

  const stylesXml = await pkg.getContent('word/styles.xml');
  const styleMaps = buildStyleMaps(stylesXml);

  const parsed = parseXml(documentXml) as XmlNode;
  const body = (parsed['w:document'] as XmlNode)?.['w:body'] as XmlNode;
  if (!body) {
    throw new Error('Invalid DOCX: missing document body');
  }

  const paragraphs: ParagraphInfo[] = [];
  const tables: TableInfo[] = [];
  const commentAnchors = new Map<string, string[]>();
  let paraIndex = 0;
  let tableIndex = 0;

  const children = body['w:p'] ? toArray(body['w:p'] as XmlNode | XmlNode[]) : [];
  const tblChildren = body['w:tbl'] ? toArray(body['w:tbl'] as XmlNode | XmlNode[]) : [];

  for (const para of children) {
    paraIndex++;
    const paragraphId = generateId('docx.p', paraIndex);
    const text = extractText(para);
    const styleId = getStyleId(para);
    const headingLevel = getHeadingLevel(styleId, styleMaps.headingLevels);
    const commentIds = extractCommentIds(para);
    const numbering = extractNumbering(para);

    paragraphs.push({
      paragraphId,
      textPreview: text.slice(0, 1200),
      styleId,
      styleName: styleId ? styleMaps.names.get(styleId) : undefined,
      headingLevel,
      isListItem: !!numbering,
      listId: numbering?.numId,
      numbering,
      format: extractParagraphFormat(para),
      runs: extractRuns(para, paragraphId).slice(0, 24),
      runCount: toArray(para['w:r'] as XmlNode | XmlNode[]).length,
      commentIds: commentIds.length > 0 ? commentIds.map((id) => `docx.comment.${id}`) : undefined,
    });

    for (const commentId of commentIds) {
      const existing = commentAnchors.get(commentId) ?? [];
      existing.push(paragraphId);
      commentAnchors.set(commentId, existing);
    }
  }

  for (const tbl of tblChildren) {
    tableIndex++;
    const rows: TableRowInfo[] = [];
    const trList = toArray(tbl['w:tr'] as XmlNode | XmlNode[]);
    const tableId = generateId('docx.tbl', tableIndex);

    for (let ri = 0; ri < trList.length; ri++) {
      const cells: TableCellInfo[] = [];
      const tcList = toArray(trList[ri]['w:tc'] as XmlNode | XmlNode[]);

      for (let ci = 0; ci < tcList.length; ci++) {
        const cellText = extractText(tcList[ci]);
        const tcPr = tcList[ci]['w:tcPr'] as XmlNode | undefined;
        const gridSpan = readNum(tcPr?.['w:gridSpan']);
        const vMerge = tcPr?.['w:vMerge'] as XmlNode | undefined;
        const vMergeVal = readVal(vMerge);
        cells.push({
          cellId: `${tableId}.r.${String(ri + 1).padStart(4, '0')}.c.${String(ci + 1).padStart(4, '0')}`,
          textPreview: cellText.slice(0, 600),
          columnSpan: gridSpan ?? 1,
          rowSpan: 1,
          width: extractMeasurement(tcPr?.['w:tcW']),
          verticalMerge: vMerge ? (vMergeVal === 'restart' ? 'restart' : 'continue') : undefined,
        });
      }

      rows.push({
        rowId: `${tableId}.r.${String(ri + 1).padStart(4, '0')}`,
        cells,
      });
    }

    tables.push({
      tableId,
      geometry: extractTableGeometry(tbl),
      rows,
    });
  }

  const headers = await extractHeadersFooters(pkg, 'header');
  const footers = await extractHeadersFooters(pkg, 'footer');

  const comments = await extractComments(pkg, commentAnchors);
  const hasComments = comments.length > 0 || paths.some((p) => p.startsWith('word/comments'));
  const hasFootnotes = paths.some((p) => p.startsWith('word/footnotes'));
  const hasEndnotes = paths.some((p) => p.startsWith('word/endnotes'));
  const hasNumbering = paths.some((p) => p.startsWith('word/numbering'));

  const relationships = extractRelationships(paths);
  const media = extractMedia(paths);

  return {
    fileType: 'docx',
    documentId: 'docx.main',
    sections: [],
    paragraphs,
    tables,
    headers,
    footers,
    comments,
    hasComments,
    hasFootnotes,
    hasEndnotes,
    hasNumbering,
    relationships,
    media,
    validation: { isValid: true, errors: [], warnings: [] },
  };
}

async function extractHeadersFooters(
  pkg: Awaited<ReturnType<typeof loadPackage>>,
  type: 'header' | 'footer'
): Promise<HeaderFooterInfo[]> {
  const results: HeaderFooterInfo[] = [];
  const paths = pkg.listPaths().filter((p) => p.startsWith(`word/${type}`));

  for (const path of paths) {
    const xml = await pkg.getContent(path);
    if (!xml) continue;

    const parsed = parseXml(xml) as XmlNode;
    const rootKey = type === 'header' ? 'w:hdr' : 'w:ftr';
    const root = parsed[rootKey] as XmlNode | undefined;
    if (!root) continue;

    const text = extractText(root);
    results.push({
      id: `docx.${type}.${path.replace(/[^0-9]/g, '') || '1'}`,
      type,
      textPreview: text.slice(0, 100),
    });
  }

  return results;
}

function extractRelationships(paths: string[]): RelationshipSummary {
  const relPaths = paths.filter((p) => p.includes('_rels') && p.endsWith('.rels'));
  let totalCount = 0;
  const types: Record<string, number> = {};

  for (const path of relPaths) {
    totalCount++;
    const type = path.includes('document') ? 'document' : 'other';
    types[type] = (types[type] ?? 0) + 1;
  }

  return { totalCount, types, externalLinks: [] };
}

function extractMedia(paths: string[]): MediaSummary {
  const mediaPaths = paths.filter((p) => p.startsWith('word/media/'));
  let images = 0;
  let audio = 0;
  let video = 0;

  for (const path of mediaPaths) {
    const ext = path.split('.').pop()?.toLowerCase() ?? '';
    if (['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg', 'webp'].includes(ext)) {
      images++;
    } else if (['mp3', 'wav', 'ogg'].includes(ext)) {
      audio++;
    } else if (['mp4', 'avi', 'mov'].includes(ext)) {
      video++;
    }
  }

  return {
    totalCount: images + audio + video,
    images,
    audio,
    video,
    totalSizeBytes: 0,
  };
}
