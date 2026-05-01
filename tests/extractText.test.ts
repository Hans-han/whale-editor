import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { composeIntent, extractReference } from '../src/utils/extractText.js';
import { createDocxWithComment } from './fixtures/createDocx.js';

describe('extractText', () => {
  it('should include the fixed input matching contract in composed intent', () => {
    const intent = composeIntent('按批注修改', []);

    assert.ok(intent.includes('Office Input Matching Contract v1'));
    assert.ok(intent.includes('DIRECT_USER_TEXT'));
    assert.ok(intent.includes('按批注修改'));
  });

  it('should include DOCX comments when extracting reference text', async () => {
    const buffer = await createDocxWithComment();
    const ref = await extractReference(buffer, 'reviewed.docx');

    assert.strictEqual(ref.kind, 'docx');
    assert.ok(ref.text.includes('Extracted DOCX Comments'));
    assert.ok(ref.text.includes('Please make this risk specific and measurable.'));
  });
});
