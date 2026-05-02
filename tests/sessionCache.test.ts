import { describe, it, beforeEach, afterEach } from 'node:test';
import assert from 'node:assert/strict';
import {
  createSession,
  getSession,
  replaceSessionDocument,
  sessionCacheStats,
  clearSessionCacheForTests,
} from '../src/utils/sessionCache.js';

describe('sessionCache limits', () => {
  beforeEach(() => {
    clearSessionCacheForTests();
  });

  afterEach(() => {
    delete process.env.SESSION_CACHE_MAX_SESSIONS;
    delete process.env.SESSION_CACHE_MAX_BYTES;
    clearSessionCacheForTests();
  });

  it('evicts the oldest sessions when the count limit is exceeded', () => {
    process.env.SESSION_CACHE_MAX_SESSIONS = '2';
    process.env.SESSION_CACHE_MAX_BYTES = '1024';

    const first = createSession(Buffer.alloc(10), 'docx', 'first.docx');
    const second = createSession(Buffer.alloc(10), 'docx', 'second.docx');
    const third = createSession(Buffer.alloc(10), 'docx', 'third.docx');

    assert.equal(getSession(first.id), null);
    assert.equal(getSession(second.id)?.id, second.id);
    assert.equal(getSession(third.id)?.id, third.id);
    assert.equal(sessionCacheStats().count, 2);
  });

  it('evicts old sessions when total cached bytes exceed the byte limit', () => {
    process.env.SESSION_CACHE_MAX_SESSIONS = '10';
    process.env.SESSION_CACHE_MAX_BYTES = '25';

    const first = createSession(Buffer.alloc(10), 'docx', 'first.docx');
    const second = createSession(Buffer.alloc(10), 'docx', 'second.docx');
    const third = createSession(Buffer.alloc(10), 'docx', 'third.docx');

    assert.equal(getSession(first.id), null);
    assert.equal(getSession(second.id)?.id, second.id);
    assert.equal(getSession(third.id)?.id, third.id);
    assert.equal(sessionCacheStats().totalBytes, 20);
  });

  it('keeps the updated session while evicting older sessions after replacement', () => {
    process.env.SESSION_CACHE_MAX_SESSIONS = '10';
    process.env.SESSION_CACHE_MAX_BYTES = '35';

    const first = createSession(Buffer.alloc(10), 'docx', 'first.docx');
    const second = createSession(Buffer.alloc(10), 'docx', 'second.docx');
    const third = createSession(Buffer.alloc(10), 'docx', 'third.docx');

    const updated = replaceSessionDocument(third.id, Buffer.alloc(30), 'docx', 'third.docx');

    assert.equal(updated?.id, third.id);
    assert.equal(getSession(first.id), null);
    assert.equal(getSession(second.id), null);
    assert.equal(getSession(third.id)?.buffer.length, 30);
    assert.equal(sessionCacheStats().totalBytes, 30);
  });
});
