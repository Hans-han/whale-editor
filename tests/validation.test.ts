import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { validateDocx } from '../src/office/docx/docxValidation.js';
import { validatePptx } from '../src/office/pptx/pptxValidation.js';
import { createMinimalDocx } from './fixtures/createDocx.js';
import { createMinimalPptx } from './fixtures/createPptx.js';

describe('validation', () => {
  it('should validate valid docx', async () => {
    const buffer = await createMinimalDocx();
    const result = await validateDocx(buffer);

    assert.strictEqual(result.isValid, true);
    assert.strictEqual(result.errors.length, 0);
  });

  it('should validate valid pptx', async () => {
    const buffer = await createMinimalPptx();
    const result = await validatePptx(buffer);

    assert.strictEqual(result.isValid, true);
    assert.strictEqual(result.errors.length, 0);
  });

  it('should reject invalid zip', async () => {
    const buffer = Buffer.from('not a zip');
    const result = await validateDocx(buffer);

    assert.strictEqual(result.isValid, false);
    assert.ok(result.errors.length > 0);
  });

  it('should detect missing required parts', async () => {
    const JSZip = (await import('jszip')).default;
    const zip = new JSZip();
    zip.file('test.txt', 'test');
    const buffer = await zip.generateAsync({ type: 'nodebuffer' });

    const result = await validateDocx(Buffer.from(buffer));
    assert.strictEqual(result.isValid, false);
  });
});
