import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { stableStringify } from '../src/utils/stableJson.js';
import { sha256, shortHash } from '../src/utils/hashing.js';
import { normalizeUploadFilename } from '../src/utils/filename.js';

describe('stableJson', () => {
  it('should produce identical output for same object', () => {
    const obj = { b: 2, a: 1, c: { z: 3, y: 2, x: 1 } };
    const result1 = stableStringify(obj);
    const result2 = stableStringify(obj);
    assert.strictEqual(result1, result2);
  });

  it('should sort keys recursively', () => {
    const obj = { c: 3, a: { z: 1, b: 2 }, b: 1 };
    const result = stableStringify(obj);
    assert.strictEqual(result, '{"a":{"b":2,"z":1},"b":1,"c":3}');
  });

  it('should handle arrays without sorting', () => {
    const obj = { items: [3, 1, 2] };
    const result = stableStringify(obj);
    assert.strictEqual(result, '{"items":[3,1,2]}');
  });

  it('should handle nested structures', () => {
    const obj = {
      level1: {
        level2: {
          level3: { z: 1, a: 2 },
        },
      },
    };
    const result = stableStringify(obj);
    assert.strictEqual(result, '{"level1":{"level2":{"level3":{"a":2,"z":1}}}}');
  });
});

describe('hashing', () => {
  it('should produce consistent sha256', () => {
    const hash1 = sha256('test');
    const hash2 = sha256('test');
    assert.strictEqual(hash1, hash2);
  });

  it('should produce 64 char hex string', () => {
    const hash = sha256('test');
    assert.strictEqual(hash.length, 64);
    assert.match(hash, /^[0-9a-f]+$/);
  });

  it('should produce different hashes for different inputs', () => {
    const hash1 = sha256('test1');
    const hash2 = sha256('test2');
    assert.notStrictEqual(hash1, hash2);
  });

  it('should produce short hash of specified length', () => {
    const hash = shortHash('test', 12);
    assert.strictEqual(hash.length, 12);
  });
});

describe('filename normalization', () => {
  it('should repair UTF-8 filenames decoded as latin1', () => {
    assert.strictEqual(
      normalizeUploadFilename('å¹´åº¦æŠ¥å‘Š.docx'),
      '年度报告.docx'
    );
  });

  it('should keep already valid Chinese filenames unchanged', () => {
    assert.strictEqual(
      normalizeUploadFilename('年度报告.docx'),
      '年度报告.docx'
    );
  });

  it('should keep ordinary ascii filenames unchanged', () => {
    assert.strictEqual(
      normalizeUploadFilename('annual-report.docx'),
      'annual-report.docx'
    );
  });
});
