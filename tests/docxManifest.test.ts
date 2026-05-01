import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { generateDocxManifest } from '../src/office/docx/docxManifest.js';
import {
  createDocxWithComment,
  createDocxWithDisabledRunProps,
  createFormattedDocx,
  createMinimalDocx,
} from './fixtures/createDocx.js';

describe('docxManifest', () => {
  it('should generate manifest from valid docx', async () => {
    const buffer = await createMinimalDocx();
    const manifest = await generateDocxManifest(buffer);

    assert.strictEqual(manifest.fileType, 'docx');
    assert.strictEqual(manifest.documentId, 'docx.main');
    assert.strictEqual(manifest.paragraphs.length, 3);
    assert.strictEqual(manifest.tables.length, 1);
  });

  it('should extract paragraph text', async () => {
    const buffer = await createMinimalDocx();
    const manifest = await generateDocxManifest(buffer);

    assert.strictEqual(manifest.paragraphs[0].textPreview, 'Hello World');
    assert.strictEqual(manifest.paragraphs[1].textPreview, 'This is a test paragraph.');
  });

  it('should generate stable IDs', async () => {
    const buffer = await createMinimalDocx();
    const manifest = await generateDocxManifest(buffer);

    assert.strictEqual(manifest.paragraphs[0].paragraphId, 'docx.p.0001');
    assert.strictEqual(manifest.paragraphs[1].paragraphId, 'docx.p.0002');
    assert.strictEqual(manifest.tables[0].tableId, 'docx.tbl.0001');
  });

  it('should extract heading level', async () => {
    const buffer = await createMinimalDocx();
    const manifest = await generateDocxManifest(buffer);

    assert.strictEqual(manifest.paragraphs[0].headingLevel, 1);
    assert.strictEqual(manifest.paragraphs[0].styleId, 'Heading1');
  });

  it('should extract table structure', async () => {
    const buffer = await createMinimalDocx();
    const manifest = await generateDocxManifest(buffer);

    const table = manifest.tables[0];
    assert.strictEqual(table.rows.length, 1);
    assert.strictEqual(table.rows[0].cells.length, 2);
    assert.strictEqual(table.rows[0].cells[0].textPreview, 'Cell 1');
  });

  it('should extract comment text and anchors', async () => {
    const buffer = await createDocxWithComment();
    const manifest = await generateDocxManifest(buffer);

    assert.strictEqual(manifest.hasComments, true);
    assert.strictEqual(manifest.comments?.length, 1);
    assert.strictEqual(manifest.comments?.[0].commentId, 'docx.comment.0');
    assert.strictEqual(manifest.comments?.[0].author, 'Reviewer');
    assert.strictEqual(
      manifest.comments?.[0].textPreview,
      'Please make this risk specific and measurable.'
    );
    assert.deepStrictEqual(manifest.comments?.[0].anchoredParagraphIds, ['docx.p.0001']);
  });

  it('should extract low-level paragraph, run, numbering, and table geometry', async () => {
    const buffer = await createFormattedDocx();
    const manifest = await generateDocxManifest(buffer);

    const para = manifest.paragraphs[0];
    assert.strictEqual(para.styleId, 'Normal');
    assert.strictEqual(para.styleName, 'Normal');
    assert.strictEqual(para.isListItem, true);
    assert.strictEqual(para.numbering?.numId, '7');
    assert.strictEqual(para.numbering?.level, '0');
    assert.strictEqual(para.format?.alignment, 'both');
    assert.strictEqual(para.format?.spacing?.after, 240);
    assert.strictEqual(para.format?.indentation?.hanging, 360);
    assert.strictEqual(para.runCount, 2);
    assert.strictEqual(para.runs?.[0].runId, 'docx.p.0001.r.0001');
    assert.strictEqual(para.runs?.[0].font, 'Aptos');
    assert.strictEqual(para.runs?.[0].bold, true);
    assert.strictEqual(para.runs?.[1].italic, true);
    assert.strictEqual(para.runs?.[1].color, 'FF0000');

    assert.strictEqual(manifest.tables[0].geometry?.width?.value, 5000);
    assert.deepStrictEqual(manifest.tables[0].geometry?.gridColumns, [2500, 2500]);
    assert.strictEqual(manifest.tables[0].rows[0].cells[0].width?.value, 2500);
  });

  it('should treat OOXML off/no run properties as false', async () => {
    const buffer = await createDocxWithDisabledRunProps();
    const manifest = await generateDocxManifest(buffer);

    assert.strictEqual(manifest.paragraphs[0].runs?.[0].bold, false);
    assert.strictEqual(manifest.paragraphs[0].runs?.[0].italic, false);
  });
});
