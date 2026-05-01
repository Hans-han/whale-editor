import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { applyPptxPatch } from '../src/office/pptx/pptxPatch.js';
import { createMinimalPptx } from './fixtures/createPptx.js';
import type { PatchEnvelope } from '../src/types/index.js';

describe('pptxPatch', () => {
  it('should replace shape text', async () => {
    const buffer = await createMinimalPptx();
    const patch: PatchEnvelope = {
      fileType: 'pptx',
      intent: 'content_edit',
      operations: [
        {
          op: 'replace_shape_text',
          targetId: 'pptx.slide.0001.shape.0002',
          payload: { text: 'New Content' },
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

    const { buffer: newBuffer, report } = await applyPptxPatch(buffer, patch);
    assert.strictEqual(report.success, true);

    const { generatePptxManifest } = await import('../src/office/pptx/pptxManifest.js');
    const manifest = await generatePptxManifest(newBuffer);
    assert.strictEqual(manifest.slides[0].shapes[1].textPreview, 'New Content');
  });

  it('should replace slide title', async () => {
    const buffer = await createMinimalPptx();
    const patch: PatchEnvelope = {
      fileType: 'pptx',
      intent: 'content_edit',
      operations: [
        {
          op: 'replace_slide_title',
          targetId: 'pptx.slide.0001',
          payload: { title: 'New Title' },
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

    const { buffer: newBuffer, report } = await applyPptxPatch(buffer, patch);
    assert.strictEqual(report.success, true);

    const { generatePptxManifest } = await import('../src/office/pptx/pptxManifest.js');
    const manifest = await generatePptxManifest(newBuffer);
    assert.strictEqual(manifest.slides[0].titleCandidate, 'New Title');
  });

  it('should reject non-pptx patch', async () => {
    const buffer = await createMinimalPptx();
    const patch = {
      fileType: 'docx',
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

    await assert.rejects(
      () => applyPptxPatch(buffer, patch as any),
      (err: Error) => {
        assert.ok(err.message.includes('fileType must be pptx'));
        return true;
      }
    );
  });
});
