import { loadPackage, parseXml } from '../ooxmlPackage.js';
import type { PackageValidation } from '../../types/index.js';

const REQUIRED_PPTX_PATHS = [
  '[Content_Types].xml',
  'ppt/presentation.xml',
  'ppt/_rels/presentation.xml.rels',
];

export async function validatePptx(buffer: Buffer): Promise<PackageValidation> {
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

  for (const path of REQUIRED_PPTX_PATHS) {
    if (!pkg.hasPath(path)) {
      errors.push(`Missing required path: ${path}`);
    }
  }

  if (errors.length > 0) {
    return { isValid: false, errors, warnings };
  }

  const presXml = await pkg.getContent('ppt/presentation.xml');
  if (presXml) {
    try {
      const parsed = parseXml(presXml);
      if (!parsed['p:presentation']) {
        errors.push('Missing p:presentation root element');
      }
    } catch (err) {
      errors.push(`Failed to parse presentation.xml: ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  const slidePaths = pkg.listPaths().filter((p) => p.match(/^ppt\/slides\/slide\d+\.xml$/));
  for (const slidePath of slidePaths) {
    const xml = await pkg.getContent(slidePath);
    if (!xml) continue;

    try {
      const parsed = parseXml(xml);
      if (!parsed['p:sld']) {
        warnings.push(`Missing p:sld root in ${slidePath}`);
      }
    } catch (err) {
      errors.push(`Failed to parse ${slidePath}: ${err instanceof Error ? err.message : String(err)}`);
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
