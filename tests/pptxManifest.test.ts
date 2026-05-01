import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { generatePptxManifest } from '../src/office/pptx/pptxManifest.js';
import { createMinimalPptx } from './fixtures/createPptx.js';

describe('pptxManifest', () => {
  it('should generate manifest from valid pptx', async () => {
    const buffer = await createMinimalPptx();
    const manifest = await generatePptxManifest(buffer);

    assert.strictEqual(manifest.fileType, 'pptx');
    assert.strictEqual(manifest.deckId, 'pptx.main');
    assert.strictEqual(manifest.slides.length, 1);
  });

  it('should extract slide shapes', async () => {
    const buffer = await createMinimalPptx();
    const manifest = await generatePptxManifest(buffer);

    const slide = manifest.slides[0];
    assert.strictEqual(slide.shapes.length, 2);
  });

  it('should extract title from shape', async () => {
    const buffer = await createMinimalPptx();
    const manifest = await generatePptxManifest(buffer);

    const slide = manifest.slides[0];
    assert.strictEqual(slide.titleCandidate, 'Slide Title');
  });

  it('should generate stable IDs', async () => {
    const buffer = await createMinimalPptx();
    const manifest = await generatePptxManifest(buffer);

    assert.strictEqual(manifest.slides[0].slideId, 'pptx.slide.0001');
    assert.strictEqual(manifest.slides[0].shapes[0].shapeId, 'pptx.shape.0001');
  });

  it('should detect placeholder types', async () => {
    const buffer = await createMinimalPptx();
    const manifest = await generatePptxManifest(buffer);

    const slide = manifest.slides[0];
    assert.strictEqual(slide.shapes[0].placeholderType, 'title');
    assert.strictEqual(slide.shapes[1].placeholderType, 'body');
  });
});
