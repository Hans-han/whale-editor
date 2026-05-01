import { PatchEnvelopeSchema, type ValidatedPatchEnvelope } from './patchTypes.js';

export class PatchValidationError extends Error {
  constructor(message: string, public errors: string[]) {
    super(message);
    this.name = 'PatchValidationError';
  }
}

export function validatePatch(data: unknown): ValidatedPatchEnvelope {
  const result = PatchEnvelopeSchema.safeParse(data);

  if (!result.success) {
    const errors = result.error.issues.map((issue) => {
      const path = issue.path.join('.');
      return path ? `${path}: ${issue.message}` : issue.message;
    });
    throw new PatchValidationError('Invalid patch envelope', errors);
  }

  return result.data;
}

export function validatePatchForFileType(data: unknown, expectedType: 'docx' | 'pptx'): ValidatedPatchEnvelope {
  const patch = validatePatch(data);

  if (patch.fileType !== expectedType) {
    throw new PatchValidationError('File type mismatch', [
      `Expected fileType "${expectedType}", got "${patch.fileType}"`,
    ]);
  }

  const validOps: Record<string, string[]> = {
    docx: [
      'replace_paragraph_text',
      'replace_text_in_paragraph',
      'update_table_cell_text',
      'insert_paragraph_after',
      'delete_paragraph',
      'apply_paragraph_style',
      'update_header_text',
      'update_footer_text',
    ],
    pptx: [
      'replace_shape_text',
      'replace_slide_title',
      'update_speaker_notes',
      'insert_slide_from_layout',
      'delete_slide',
      'move_shape',
      'resize_shape',
      'apply_text_style',
      'replace_image',
      'fit_text_to_shape',
    ],
  };

  const allowed = validOps[expectedType];
  for (const op of patch.operations) {
    if (!allowed.includes(op.op)) {
      throw new PatchValidationError('Invalid operation for file type', [
        `Operation "${op.op}" is not valid for ${expectedType}`,
      ]);
    }
  }

  return patch;
}
