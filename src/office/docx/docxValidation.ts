import { loadPackage, parseXml } from '../ooxmlPackage.js';
import type { PackageValidation } from '../../types/index.js';

const REQUIRED_DOCX_PATHS = [
  '[Content_Types].xml',
  'word/document.xml',
  'word/_rels/document.xml.rels',
];

export async function validateDocx(buffer: Buffer): Promise<PackageValidation> {
  const errors: string[] = [];
  const warnings: string[] = [];

  let pkg;
  try {
    pkg = await loadPackage(buffer);
  } catch (err) {
    return {
      isValid: false,
      errors: [`Failed to open ZIP: ${err instanceof Error ? err.message : String(err)}`],
      warnings: [],
    };
  }

  for (const path of REQUIRED_DOCX_PATHS) {
    if (!pkg.hasPath(path)) {
      errors.push(`Missing required path: ${path}`);
    }
  }

  if (errors.length > 0) {
    return { isValid: false, errors, warnings };
  }

  const documentXml = await pkg.getContent('word/document.xml');
  if (documentXml) {
    try {
      const parsed = parseXml(documentXml);
      if (!parsed['w:document']) {
        errors.push('Missing w:document root element');
      }
    } catch (err) {
      errors.push(`Failed to parse document.xml: ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  const relsXml = await pkg.getContent('word/_rels/document.xml.rels');
  if (relsXml) {
    try {
      parseXml(relsXml);
    } catch (err) {
      warnings.push(`Failed to parse document.xml.rels: ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  const contentTypes = await pkg.getContent('[Content_Types].xml');
  if (contentTypes) {
    try {
      parseXml(contentTypes);
    } catch (err) {
      errors.push(`Failed to parse [Content_Types].xml: ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  return {
    isValid: errors.length === 0,
    errors,
    warnings,
  };
}
