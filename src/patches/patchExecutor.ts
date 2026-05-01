import type { PatchEnvelope, ExecutionReport } from '../types/index.js';
import { validatePatchForFileType } from './patchValidator.js';
import { applyDocxPatch } from '../office/docx/docxPatch.js';
import { applyPptxPatch } from '../office/pptx/pptxPatch.js';

export interface PatchExecutionResult {
  buffer: Buffer;
  report: ExecutionReport;
}

export async function executePatch(
  buffer: Buffer,
  patchData: unknown
): Promise<PatchExecutionResult> {
  const patch = validatePatchForFileType(patchData, detectFileType(patchData));

  if (patch.fileType === 'docx') {
    return applyDocxPatch(buffer, patch);
  }

  return applyPptxPatch(buffer, patch);
}

function detectFileType(data: unknown): 'docx' | 'pptx' {
  if (data && typeof data === 'object' && 'fileType' in data) {
    const ft = (data as Record<string, unknown>).fileType;
    if (ft === 'docx' || ft === 'pptx') return ft;
  }
  throw new Error('Cannot detect fileType from patch data');
}
