import JSZip from 'jszip';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

export interface OoxmlPackage {
  zip: JSZip;
  getContent(path: string): Promise<string | null>;
  setContent(path: string, content: string): Promise<void>;
  getBuffer(): Promise<Buffer>;
  listPaths(): string[];
  hasPath(path: string): boolean;
}

const XML_PARSER_OPTIONS = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  allowBooleanAttributes: true,
  parseAttributeValue: true,
  trimValues: false,
};

const XML_BUILDER_OPTIONS = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  format: false,
  suppressEmptyNode: true,
  suppressBooleanAttributes: false,
};

export function createXmlParser(): XMLParser {
  return new XMLParser(XML_PARSER_OPTIONS);
}

export function createXmlBuilder(): XMLBuilder {
  return new XMLBuilder(XML_BUILDER_OPTIONS);
}

export async function loadPackage(buffer: Buffer): Promise<OoxmlPackage> {
  const zip = await JSZip.loadAsync(buffer);
  const paths = Object.keys(zip.files);

  return {
    zip,
    async getContent(path: string): Promise<string | null> {
      const file = zip.file(path);
      if (!file) return null;
      return file.async('string');
    },
    async setContent(path: string, content: string): Promise<void> {
      zip.file(path, content);
    },
    async getBuffer(): Promise<Buffer> {
      const buf = await zip.generateAsync({ type: 'nodebuffer' });
      return Buffer.from(buf);
    },
    listPaths(): string[] {
      return paths;
    },
    hasPath(path: string): boolean {
      return zip.file(path) !== null;
    },
  };
}

export function parseXml(xml: string): Record<string, unknown> {
  const parser = createXmlParser();
  return parser.parse(xml) as Record<string, unknown>;
}

export function buildXml(obj: Record<string, unknown>): string {
  const builder = createXmlBuilder();
  return builder.build(obj);
}

export function getNestedValue(obj: unknown, path: string): unknown {
  const parts = path.split('.');
  let current = obj;
  for (const part of parts) {
    if (current === null || current === undefined) return undefined;
    if (typeof current !== 'object') return undefined;
    current = (current as Record<string, unknown>)[part];
  }
  return current;
}

export function setNestedValue(obj: Record<string, unknown>, path: string, value: unknown): void {
  const parts = path.split('.');
  let current = obj;
  for (let i = 0; i < parts.length - 1; i++) {
    const part = parts[i];
    if (!(part in current) || typeof current[part] !== 'object') {
      current[part] = {};
    }
    current = current[part] as Record<string, unknown>;
  }
  current[parts[parts.length - 1]] = value;
}
