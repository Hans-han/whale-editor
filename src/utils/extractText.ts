import mammoth from 'mammoth';
import { extractText as unpdfExtract } from 'unpdf';
import WordExtractor from 'word-extractor';
import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { convertToOoxml } from './convertOoxml.js';
import { logger } from './logger.js';

export type ReferenceKind = 'pdf' | 'docx' | 'doc' | 'pptx' | 'ppt' | 'md' | 'txt';

const wordExtractor = new WordExtractor();
const INPUT_MATCHER_PROMPT_PATH = resolve(import.meta.dirname, '../../prompts/office-input-matcher.md');
let cachedInputMatcherPrompt: string | null = null;

export interface ExtractedReference {
  filename: string;
  kind: ReferenceKind;
  text: string;
}

const MAX_TEXT_CHARS = 100_000; // per reference, to keep prompts bounded

function getInputMatcherPrompt(): string {
  if (cachedInputMatcherPrompt) return cachedInputMatcherPrompt;
  cachedInputMatcherPrompt = readFileSync(INPUT_MATCHER_PROMPT_PATH, 'utf8');
  return cachedInputMatcherPrompt;
}

function isPdf(b: Buffer): boolean {
  return (
    b.length >= 4 &&
    b[0] === 0x25 &&
    b[1] === 0x50 &&
    b[2] === 0x44 &&
    b[3] === 0x46
  );
}

function isZip(b: Buffer): boolean {
  return b.length >= 2 && b[0] === 0x50 && b[1] === 0x4b;
}

// OLE compound file (legacy .doc): D0 CF 11 E0 A1 B1 1A E1
function isOleCompound(b: Buffer): boolean {
  return (
    b.length >= 8 &&
    b[0] === 0xd0 &&
    b[1] === 0xcf &&
    b[2] === 0x11 &&
    b[3] === 0xe0 &&
    b[4] === 0xa1 &&
    b[5] === 0xb1 &&
    b[6] === 0x1a &&
    b[7] === 0xe1
  );
}

function detectKind(filename: string, buf: Buffer): ReferenceKind {
  const lower = filename.toLowerCase();
  if (lower.endsWith('.pdf') || isPdf(buf)) return 'pdf';
  if (lower.endsWith('.docx') && isZip(buf)) return 'docx';
  if (lower.endsWith('.pptx') && isZip(buf)) return 'pptx';
  if (lower.endsWith('.ppt')) return 'ppt';
  if (lower.endsWith('.doc') || isOleCompound(buf)) return 'doc';
  if (lower.endsWith('.md') || lower.endsWith('.markdown')) return 'md';
  if (lower.endsWith('.docx')) return 'docx';
  if (lower.endsWith('.pptx')) return 'pptx';
  return 'txt';
}

async function extractPptxText(buffer: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(buffer);
  const slidePaths = Object.keys(zip.files)
    .filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p))
    .sort((a, b) => {
      const an = parseInt(a.match(/(\d+)/)?.[1] ?? '0', 10);
      const bn = parseInt(b.match(/(\d+)/)?.[1] ?? '0', 10);
      return an - bn;
    });

  const parser = new XMLParser({
    ignoreAttributes: true,
    preserveOrder: false,
    textNodeName: '#text',
    isArray: () => false,
  });

  const slideTexts: string[] = [];
  for (let i = 0; i < slidePaths.length; i++) {
    const xml = await zip.files[slidePaths[i]].async('string');
    const obj = parser.parse(xml);
    const texts: string[] = [];
    collectAText(obj, texts);
    if (texts.length > 0) {
      slideTexts.push(`### Slide ${i + 1}\n` + texts.join('\n'));
    }
  }
  return slideTexts.join('\n\n');
}

async function extractDocxText(buffer: Buffer): Promise<string> {
  const raw = await mammoth.extractRawText({ buffer });
  const body = (raw.value ?? '').trim();
  const comments = await extractDocxComments(buffer);

  if (comments.length === 0) return body;

  return [
    body,
    '## Extracted DOCX Comments',
    ...comments.map((c, i) => {
      const meta = [
        `comment ${i + 1}`,
        c.author ? `author=${c.author}` : '',
        c.date ? `date=${c.date}` : '',
      ].filter(Boolean).join(' · ');
      return `### ${meta}\n${c.text}`;
    }),
  ].filter(Boolean).join('\n\n');
}

async function extractDocxComments(buffer: Buffer): Promise<Array<{ author?: string; date?: string; text: string }>> {
  if (!isZip(buffer)) return [];
  const zip = await JSZip.loadAsync(buffer);
  const commentsFile = zip.file('word/comments.xml');
  if (!commentsFile) return [];

  const xml = await commentsFile.async('string');
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    trimValues: true,
  });
  const parsed = parser.parse(xml) as Record<string, unknown>;
  const root = parsed['w:comments'] as Record<string, unknown> | undefined;
  if (!root) return [];

  const raw = root['w:comment'];
  const comments = Array.isArray(raw) ? raw : raw ? [raw] : [];
  return comments
    .map((item): { author?: string; date?: string; text: string } | null => {
      if (!item || typeof item !== 'object') return null;
      const obj = item as Record<string, unknown>;
      const texts: string[] = [];
      collectWText(obj, texts);
      const text = texts.join('').trim();
      if (!text) return null;
      return {
        author: typeof obj['@_w:author'] === 'string' ? obj['@_w:author'] : undefined,
        date: obj['@_w:date'] !== undefined ? String(obj['@_w:date']) : undefined,
        text,
      };
    })
    .filter((c): c is { author?: string; date?: string; text: string } => c !== null);
}

// Walk parsed XML, collecting any node under key 'a:t' (text run content).
function collectAText(node: unknown, out: string[]): void {
  if (node === null || node === undefined) return;
  if (typeof node === 'string') return;
  if (Array.isArray(node)) {
    for (const item of node) collectAText(item, out);
    return;
  }
  if (typeof node === 'object') {
    for (const [k, v] of Object.entries(node as Record<string, unknown>)) {
      if (k === 'a:t') {
        if (typeof v === 'string' && v.trim()) out.push(v);
        else if (typeof v === 'number') out.push(String(v));
        else if (Array.isArray(v)) {
          for (const t of v) if (typeof t === 'string' && t.trim()) out.push(t);
        }
      } else {
        collectAText(v, out);
      }
    }
  }
}

function collectWText(node: unknown, out: string[]): void {
  if (node === null || node === undefined) return;
  if (typeof node === 'string') return;
  if (Array.isArray(node)) {
    for (const item of node) collectWText(item, out);
    return;
  }
  if (typeof node === 'object') {
    for (const [k, v] of Object.entries(node as Record<string, unknown>)) {
      if (k === 'w:t') {
        if (typeof v === 'string') out.push(v);
        else if (v && typeof v === 'object' && '#text' in (v as Record<string, unknown>)) {
          out.push(String((v as Record<string, unknown>)['#text']));
        }
      } else if (!k.startsWith('@_')) {
        collectWText(v, out);
      }
    }
  }
}

function clamp(text: string): string {
  if (text.length <= MAX_TEXT_CHARS) return text;
  return text.slice(0, MAX_TEXT_CHARS) + `\n\n[...truncated, ${text.length - MAX_TEXT_CHARS} chars omitted]`;
}

export async function extractReference(
  buffer: Buffer,
  filename: string
): Promise<ExtractedReference> {
  const kind = detectKind(filename, buffer);

  try {
    if (kind === 'pdf') {
      // unpdf accepts Uint8Array
      const { text } = await unpdfExtract(new Uint8Array(buffer));
      const joined = Array.isArray(text) ? text.join('\n\n') : (text ?? '');
      return { filename, kind, text: clamp(joined.trim()) };
    }
    if (kind === 'docx') {
      const text = await extractDocxText(buffer);
      return { filename, kind, text: clamp(text.trim()) };
    }
    if (kind === 'doc') {
      const extracted = await wordExtractor.extract(buffer);
      const body = (extracted.getBody?.() ?? '').toString();
      return { filename, kind, text: clamp(body.trim()) };
    }
    if (kind === 'pptx') {
      const text = await extractPptxText(buffer);
      return { filename, kind, text: clamp(text.trim()) };
    }
    if (kind === 'ppt') {
      // Convert via libreoffice if available, then extract from the resulting pptx
      const converted = await convertToOoxml(buffer, filename);
      const text = await extractPptxText(converted.buffer);
      return { filename, kind, text: clamp(text.trim()) };
    }
    // md, txt
    return { filename, kind, text: clamp(buffer.toString('utf-8').trim()) };
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    logger.warn('Reference extraction failed', { filename, kind, error: msg });
    throw new Error(`无法解析参考文件 ${filename}（${kind}）：${msg}`);
  }
}

export function composeIntent(
  freeText: string,
  references: ExtractedReference[]
): string {
  const parts: string[] = [];
  parts.push('## Fixed Input Matching Contract\n\n' + getInputMatcherPrompt().trim());

  if (freeText.trim()) {
    parts.push('## Input Block 1 · DIRECT_USER_TEXT\n\n' + freeText.trim());
  } else {
    parts.push('## Input Block 1 · DIRECT_USER_TEXT\n\n(无直接文字指令；请根据参考材料或文档批注进行最小安全匹配)');
  }

  references.forEach((ref, i) => {
    parts.push(
      `## Input Block ${i + 2} · REFERENCE_FILE · ${ref.kind} · ${ref.filename}\n\n${ref.text || '(空)'}`
    );
  });

  return parts.join('\n\n---\n\n');
}
