import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { applyDocxPatch } from '../src/office/docx/docxPatch.js';
import { createFormattedDocx, createMinimalDocx } from './fixtures/createDocx.js';
import type { PatchEnvelope } from '../src/types/index.js';

describe('docxPatch', () => {
  it('should replace paragraph text', async () => {
    const buffer = await createMinimalDocx();
    const patch: PatchEnvelope = {
      fileType: 'docx',
      intent: 'content_edit',
      operations: [
        {
          op: 'replace_paragraph_text',
          targetId: 'docx.p.0001',
          payload: { text: 'New Title' },
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

    const { buffer: newBuffer, report } = await applyDocxPatch(buffer, patch);
    assert.strictEqual(report.success, true);
    assert.strictEqual(report.operationsExecuted, 1);

    const { generateDocxManifest } = await import('../src/office/docx/docxManifest.js');
    const manifest = await generateDocxManifest(newBuffer);
    assert.strictEqual(manifest.paragraphs[0].textPreview, 'New Title');
  });

  it('should update table cell', async () => {
    const buffer = await createMinimalDocx();
    const patch: PatchEnvelope = {
      fileType: 'docx',
      intent: 'content_edit',
      operations: [
        {
          op: 'update_table_cell_text',
          targetId: 'docx.tbl.0001.r.0001.c.0001',
          payload: { text: 'Updated Cell' },
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

    const { buffer: newBuffer, report } = await applyDocxPatch(buffer, patch);
    assert.strictEqual(report.success, true);

    const { generateDocxManifest } = await import('../src/office/docx/docxManifest.js');
    const manifest = await generateDocxManifest(newBuffer);
    assert.strictEqual(manifest.tables[0].rows[0].cells[0].textPreview, 'Updated Cell');
  });

  it('should replace text inside existing runs without losing run formatting', async () => {
    const buffer = await createFormattedDocx();
    const patch: PatchEnvelope = {
      fileType: 'docx',
      intent: 'content_edit',
      operations: [
        {
          op: 'replace_text_in_paragraph',
          targetId: 'docx.p.0001',
          payload: { find: 'old', replace: 'new' },
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

    const { buffer: newBuffer, report } = await applyDocxPatch(buffer, patch);
    assert.strictEqual(report.success, true);

    const { generateDocxManifest } = await import('../src/office/docx/docxManifest.js');
    const manifest = await generateDocxManifest(newBuffer);
    assert.strictEqual(manifest.paragraphs[0].textPreview, 'The new value');
    assert.strictEqual(manifest.paragraphs[0].runCount, 2);
    assert.strictEqual(manifest.paragraphs[0].runs?.[1].textPreview, 'new value');
    assert.strictEqual(manifest.paragraphs[0].runs?.[1].italic, true);
    assert.strictEqual(manifest.paragraphs[0].runs?.[1].color, 'FF0000');
    assert.strictEqual(manifest.paragraphs[0].numbering?.numId, '7');
  });

  it('should preserve xml:space when a replacement introduces edge spaces', async () => {
    const buffer = await createFormattedDocx();
    const patch: PatchEnvelope = {
      fileType: 'docx',
      intent: 'content_edit',
      operations: [
        {
          op: 'replace_text_in_paragraph',
          targetId: 'docx.p.0001',
          payload: { find: 'old', replace: ' new ' },
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

    const { buffer: newBuffer, report } = await applyDocxPatch(buffer, patch);
    assert.strictEqual(report.success, true);

    const zip = await JSZip.loadAsync(newBuffer);
    const xml = await zip.file('word/document.xml')?.async('string');
    assert.match(xml ?? '', /<w:t xml:space="preserve"> new  value<\/w:t>/);
  });

  it('should delete paragraph', async () => {
    const buffer = await createMinimalDocx();
    const patch: PatchEnvelope = {
      fileType: 'docx',
      intent: 'content_edit',
      operations: [
        {
          op: 'delete_paragraph',
          targetId: 'docx.p.0003',
          payload: {},
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

    const { buffer: newBuffer, report } = await applyDocxPatch(buffer, patch);
    assert.strictEqual(report.success, true);

    const { generateDocxManifest } = await import('../src/office/docx/docxManifest.js');
    const manifest = await generateDocxManifest(newBuffer);
    assert.strictEqual(manifest.paragraphs.length, 2);
  });

  it('should reject non-docx patch', async () => {
    const buffer = await createMinimalDocx();
    const patch = {
      fileType: 'pptx',
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
      () => applyDocxPatch(buffer, patch as any),
      (err: Error) => {
        assert.ok(err.message.includes('fileType must be docx'));
        return true;
      }
    );
  });
});
