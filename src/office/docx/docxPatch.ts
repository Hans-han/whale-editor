import { loadPackage, parseXml, buildXml } from '../ooxmlPackage.js';
import type {
  PatchEnvelope,
  PatchOperation,
  ExecutionReport,
  ExecutionError,
} from '../../types/index.js';

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

function findParagraphByIndex(body: Record<string, unknown>, index: number): Record<string, unknown> | null {
  const paras = toArray(body['w:p'] as Record<string, unknown> | Record<string, unknown>[]);
  if (index < 0 || index >= paras.length) return null;
  return paras[index];
}

function findTableByIndex(body: Record<string, unknown>, index: number): Record<string, unknown> | null {
  const tables = toArray(body['w:tbl'] as Record<string, unknown> | Record<string, unknown>[]);
  if (index < 0 || index >= tables.length) return null;
  return tables[index];
}

interface TextNodeRef {
  node: Record<string, unknown>;
  key: string;
  text: string;
  start: number;
  end: number;
}

function parseTargetId(targetId: string): { type: string; indices: number[] } {
  const parts = targetId.split('.');
  const type = parts[0] === 'docx' ? parts[1] : parts[0];
  const indices: number[] = [];

  for (let i = 2; i < parts.length; i++) {
    const num = parseInt(parts[i], 10);
    if (!isNaN(num)) indices.push(num);
  }

  return { type, indices };
}

export async function applyDocxPatch(
  buffer: Buffer,
  patch: PatchEnvelope
): Promise<{ buffer: Buffer; report: ExecutionReport }> {
  if (patch.fileType !== 'docx') {
    throw new Error('Invalid patch: fileType must be docx');
  }

  const pkg = await loadPackage(buffer);
  const documentXml = await pkg.getContent('word/document.xml');

  if (!documentXml) {
    throw new Error('Invalid DOCX: missing word/document.xml');
  }

  const parsed = parseXml(documentXml) as Record<string, unknown>;
  const doc = parsed['w:document'] as Record<string, unknown>;
  const body = doc['w:body'] as Record<string, unknown>;

  const errors: ExecutionError[] = [];
  const warnings: string[] = [];
  let executed = 0;
  let failed = 0;

  for (let i = 0; i < patch.operations.length; i++) {
    const op = patch.operations[i];
    try {
      applyOperation(body, op, warnings);
      executed++;
    } catch (err) {
      failed++;
      errors.push({
        operationIndex: i,
        operation: op.op,
        targetId: op.targetId,
        error: err instanceof Error ? err.message : String(err),
        recoverable: false,
      });
    }
  }

  const newXml = buildXml(parsed);
  await pkg.setContent('word/document.xml', newXml);
  const newBuffer = await pkg.getBuffer();

  return {
    buffer: newBuffer,
    report: {
      success: failed === 0,
      operationsExecuted: executed,
      operationsFailed: failed,
      errors,
      warnings,
    },
  };
}

function applyOperation(body: Record<string, unknown>, op: PatchOperation, warnings: string[]): void {
  const { type, indices } = parseTargetId(op.targetId);

  switch (op.op) {
    case 'replace_paragraph_text': {
      const para = findParagraphByIndex(body, indices[0] - 1);
      if (!para) throw new Error(`Paragraph not found: ${op.targetId}`);
      replaceParagraphContent(para, op.payload.text);
      break;
    }

    case 'replace_text_in_paragraph': {
      const para = findParagraphByIndex(body, indices[0] - 1);
      if (!para) throw new Error(`Paragraph not found: ${op.targetId}`);
      const result = replaceTextInExistingTextNodes(para, op.payload.find, op.payload.replace);
      if (!result.replaced) {
        throw new Error(`Text "${op.payload.find}" not found in paragraph`);
      }
      if (result.crossedTextNodes) {
        warnings.push(
          `replace_text_in_paragraph crossed multiple w:t nodes for ${op.targetId}; replacement text follows existing run boundaries as closely as possible.`
        );
      }
      break;
    }

    case 'update_table_cell_text': {
      const tblIdx = indices[0] - 1;
      const rowIdx = indices[1] - 1;
      const cellIdx = indices[2] - 1;
      const table = findTableByIndex(body, tblIdx);
      if (!table) throw new Error(`Table not found: ${op.targetId}`);
      const rows = toArray(table['w:tr'] as Record<string, unknown> | Record<string, unknown>[]);
      if (rowIdx < 0 || rowIdx >= rows.length) throw new Error(`Row not found: ${op.targetId}`);
      const cells = toArray(rows[rowIdx]['w:tc'] as Record<string, unknown> | Record<string, unknown>[]);
      if (cellIdx < 0 || cellIdx >= cells.length) throw new Error(`Cell not found: ${op.targetId}`);
      replaceCellContent(cells[cellIdx], op.payload.text);
      break;
    }

    case 'insert_paragraph_after': {
      const para = findParagraphByIndex(body, indices[0] - 1);
      if (!para) throw new Error(`Paragraph not found: ${op.targetId}`);
      const newPara = createParagraph(op.payload.text, op.payload.styleId);
      const paras = toArray(body['w:p'] as Record<string, unknown> | Record<string, unknown>[]);
      const idx = paras.indexOf(para);
      paras.splice(idx + 1, 0, newPara);
      body['w:p'] = paras;
      break;
    }

    case 'delete_paragraph': {
      const paras = toArray(body['w:p'] as Record<string, unknown> | Record<string, unknown>[]);
      const idx = indices[0] - 1;
      if (idx < 0 || idx >= paras.length) throw new Error(`Paragraph not found: ${op.targetId}`);
      paras.splice(idx, 1);
      body['w:p'] = paras;
      break;
    }

    case 'apply_paragraph_style': {
      const para = findParagraphByIndex(body, indices[0] - 1);
      if (!para) throw new Error(`Paragraph not found: ${op.targetId}`);
      if (!para['w:pPr']) para['w:pPr'] = {};
      (para['w:pPr'] as Record<string, unknown>)['w:pStyle'] = { '@_w:val': op.payload.styleId };
      break;
    }

    default:
      throw new Error(`Unsupported operation: ${op.op}`);
  }
}

function replaceParagraphContent(para: Record<string, unknown>, text: string): void {
  const runs = toArray(para['w:r'] as Record<string, unknown> | Record<string, unknown>[]);

  if (runs.length > 0) {
    const firstRun = runs[0];
    firstRun['w:t'] = { '#text': text, '@_xml:space': 'preserve' };

    for (let i = 1; i < runs.length; i++) {
      delete runs[i]['w:t'];
    }
  } else {
    para['w:r'] = {
      'w:t': { '#text': text, '@_xml:space': 'preserve' },
    };
  }
}

function getTextValue(t: unknown): string {
  if (typeof t === 'string') return t;
  if (t && typeof t === 'object' && '#text' in (t as Record<string, unknown>)) {
    return String((t as Record<string, unknown>)['#text'] ?? '');
  }
  return '';
}

function setTextValue(node: Record<string, unknown>, key: string, value: string): void {
  const current = node[key];
  if (typeof current === 'string') {
    node[key] = { '#text': value, '@_xml:space': 'preserve' };
    return;
  }
  if (current && typeof current === 'object') {
    (current as Record<string, unknown>)['#text'] = value;
    (current as Record<string, unknown>)['@_xml:space'] = 'preserve';
    return;
  }
  node[key] = { '#text': value, '@_xml:space': 'preserve' };
}

function collectTextNodes(node: unknown, refs: TextNodeRef[], cursor: { value: number }): void {
  if (!node || typeof node !== 'object') return;

  if (Array.isArray(node)) {
    for (const item of node) collectTextNodes(item, refs, cursor);
    return;
  }

  const obj = node as Record<string, unknown>;
  if ('w:t' in obj) {
    const t = obj['w:t'];
    if (Array.isArray(t)) {
      for (let i = 0; i < t.length; i++) {
        const text = getTextValue(t[i]);
        refs.push({ node: t as unknown as Record<string, unknown>, key: String(i), text, start: cursor.value, end: cursor.value + text.length });
        cursor.value += text.length;
      }
      return;
    }
    const text = getTextValue(t);
    refs.push({ node: obj, key: 'w:t', text, start: cursor.value, end: cursor.value + text.length });
    cursor.value += text.length;
    return;
  }

  for (const key of Object.keys(obj)) {
    if (key.startsWith('@_')) continue;
    collectTextNodes(obj[key], refs, cursor);
  }
}

function replaceTextInExistingTextNodes(
  para: Record<string, unknown>,
  find: string,
  replace: string
): { replaced: boolean; crossedTextNodes: boolean } {
  const refs: TextNodeRef[] = [];
  collectTextNodes(para, refs, { value: 0 });
  const fullText = refs.map((r) => r.text).join('');
  const start = fullText.indexOf(find);
  if (start === -1) return { replaced: false, crossedTextNodes: false };

  const end = start + find.length;
  const touched = refs.filter((r) => r.end > start && r.start < end);
  if (touched.length === 0) return { replaced: false, crossedTextNodes: false };

  const first = touched[0];
  const last = touched[touched.length - 1];
  const prefix = first.text.slice(0, Math.max(0, start - first.start));
  const suffix = last.text.slice(Math.max(0, end - last.start));

  if (touched.length === 1) {
    setTextValue(first.node, first.key, prefix + replace + suffix);
    return { replaced: true, crossedTextNodes: false };
  }

  setTextValue(first.node, first.key, prefix + replace);
  for (let i = 1; i < touched.length - 1; i++) {
    setTextValue(touched[i].node, touched[i].key, '');
  }
  setTextValue(last.node, last.key, suffix);
  return { replaced: true, crossedTextNodes: true };
}

function replaceCellContent(cell: Record<string, unknown>, text: string): void {
  const paras = toArray(cell['w:p'] as Record<string, unknown> | Record<string, unknown>[]);
  if (paras.length > 0) {
    replaceParagraphContent(paras[0], text);
    for (let i = 1; i < paras.length; i++) {
      paras.splice(i, 1);
    }
    cell['w:p'] = paras;
  } else {
    cell['w:p'] = createParagraph(text);
  }
}

function createParagraph(text: string, styleId?: string): Record<string, unknown> {
  const para: Record<string, unknown> = {
    'w:r': {
      'w:t': { '#text': text, '@_xml:space': 'preserve' },
    },
  };

  if (styleId) {
    para['w:pPr'] = {
      'w:pStyle': { '@_w:val': styleId },
    };
  }

  return para;
}
