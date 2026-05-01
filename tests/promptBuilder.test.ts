import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { buildPrompt } from '../src/agent/promptBuilder.js';
import type { OfficeManifest } from '../src/types/index.js';

const mockManifest: OfficeManifest = {
  fileType: 'docx',
  documentId: 'docx.main',
  sections: [],
  paragraphs: [
    {
      paragraphId: 'docx.p.0001',
      textPreview: 'Hello World',
      isListItem: false,
      headingLevel: 1,
    },
  ],
  tables: [],
  headers: [],
  footers: [],
  hasComments: false,
  hasFootnotes: false,
  hasEndnotes: false,
  hasNumbering: false,
  relationships: { totalCount: 0, types: {}, externalLinks: [] },
  media: { totalCount: 0, images: 0, audio: 0, video: 0, totalSizeBytes: 0 },
  validation: { isValid: true, errors: [], warnings: [] },
};

describe('promptBuilder', () => {
  it('should produce deterministic output', () => {
    const result1 = buildPrompt({
      systemRules: 'test rules',
      toolSpecs: 'test tools',
      patchSchema: 'test schema',
      validationRules: 'test validation',
      manifest: mockManifest,
      userTask: 'test task',
    });

    const result2 = buildPrompt({
      systemRules: 'test rules',
      toolSpecs: 'test tools',
      patchSchema: 'test schema',
      validationRules: 'test validation',
      manifest: mockManifest,
      userTask: 'test task',
    });

    assert.strictEqual(result1.stablePrefixHash, result2.stablePrefixHash);
  });

  it('should produce different hashes for different inputs', () => {
    const result1 = buildPrompt({
      systemRules: 'rules1',
      toolSpecs: 'tools',
      patchSchema: 'schema',
      validationRules: 'validation',
      manifest: mockManifest,
      userTask: 'task1',
    });

    const result2 = buildPrompt({
      systemRules: 'rules2',
      toolSpecs: 'tools',
      patchSchema: 'schema',
      validationRules: 'validation',
      manifest: mockManifest,
      userTask: 'task2',
    });

    assert.notStrictEqual(result1.stablePrefixHash, result2.stablePrefixHash);
  });

  it('should generate messages with system and user roles', () => {
    const result = buildPrompt({
      systemRules: 'rules',
      toolSpecs: 'tools',
      patchSchema: 'schema',
      validationRules: 'validation',
      manifest: mockManifest,
      userTask: 'task',
    });

    assert.strictEqual(result.messages.length, 2);
    assert.strictEqual(result.messages[0].role, 'system');
    assert.strictEqual(result.messages[1].role, 'user');
  });

  it('should include manifest in user message', () => {
    const result = buildPrompt({
      systemRules: 'rules',
      toolSpecs: 'tools',
      patchSchema: 'schema',
      validationRules: 'validation',
      manifest: mockManifest,
      userTask: 'my task',
    });

    assert.ok(result.messages[1].content.includes('my task'));
    assert.ok(result.messages[1].content.includes('docx.main'));
  });

  it('should initialize cache metrics', () => {
    const result = buildPrompt({
      systemRules: 'rules',
      toolSpecs: 'tools',
      patchSchema: 'schema',
      validationRules: 'validation',
      manifest: mockManifest,
      userTask: 'task',
    });

    assert.strictEqual(result.cacheMetrics.stablePrefixHash, result.stablePrefixHash);
    assert.strictEqual(result.cacheMetrics.cacheHitTokens, 0);
    assert.strictEqual(result.cacheMetrics.cacheMissTokens, 0);
    assert.strictEqual(result.cacheMetrics.cacheHitRatio, 0);
  });

  it('should include Office-native formatting rules in the default system prompt', () => {
    const result = buildPrompt({
      systemRules: '',
      toolSpecs: 'tools',
      patchSchema: 'schema',
      validationRules: 'validation',
      manifest: mockManifest,
      userTask: 'task',
    });

    assert.ok(result.messages[0].content.includes('Word-Native Formatting Contract'));
    assert.ok(result.messages[0].content.includes('Numbering fidelity'));
    assert.ok(result.messages[0].content.includes('Table fidelity'));
  });
});
