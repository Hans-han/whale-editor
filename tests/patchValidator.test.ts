import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { validatePatch, PatchValidationError } from '../src/patches/patchValidator.js';

describe('patchValidator', () => {
  it('should validate a valid docx patch', () => {
    const patch = {
      fileType: 'docx',
      intent: 'content_edit',
      operations: [
        {
          op: 'replace_paragraph_text',
          targetId: 'docx.p.0001',
          payload: { text: 'New text' },
          preserveFormatting: true,
        },
      ],
      preserve: {
        styles: true,
        themes: true,
        relationships: true,
        comments: true,
        trackedChanges: true,
        speakerNotes: true,
      },
    };

    const result = validatePatch(patch);
    assert.strictEqual(result.fileType, 'docx');
    assert.strictEqual(result.operations.length, 1);
  });

  it('should validate a valid pptx patch', () => {
    const patch = {
      fileType: 'pptx',
      intent: 'content_edit',
      operations: [
        {
          op: 'replace_shape_text',
          targetId: 'pptx.slide.0001.shape.0001',
          payload: { text: 'New text' },
          preserveFormatting: true,
        },
      ],
      preserve: {
        styles: true,
        themes: true,
        relationships: true,
        comments: true,
        trackedChanges: true,
        speakerNotes: true,
      },
    };

    const result = validatePatch(patch);
    assert.strictEqual(result.fileType, 'pptx');
  });

  it('should reject invalid fileType', () => {
    const patch = {
      fileType: 'xlsx',
      intent: 'content_edit',
      operations: [],
      preserve: {
        styles: true,
        themes: true,
        relationships: true,
        comments: true,
        trackedChanges: true,
        speakerNotes: true,
      },
    };

    assert.throws(() => validatePatch(patch), PatchValidationError);
  });

  it('should reject missing required fields', () => {
    const patch = {
      fileType: 'docx',
    };

    assert.throws(() => validatePatch(patch), PatchValidationError);
  });

  it('should reject too many operations', () => {
    const operations = Array.from({ length: 21 }, (_, i) => ({
      op: 'replace_paragraph_text',
      targetId: `docx.p.${String(i + 1).padStart(4, '0')}`,
      payload: { text: `Text ${i}` },
      preserveFormatting: true,
    }));

    const patch = {
      fileType: 'docx',
      intent: 'content_edit',
      operations,
      preserve: {
        styles: true,
        themes: true,
        relationships: true,
        comments: true,
        trackedChanges: true,
        speakerNotes: true,
      },
    };

    assert.throws(() => validatePatch(patch), PatchValidationError);
  });
});
