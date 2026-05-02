import { describe, it, before, beforeEach, after, afterEach } from 'node:test';
import assert from 'node:assert/strict';
import { request } from 'node:http';
import type { Server } from 'node:http';
import { app } from '../src/api/server.js';
import {
  clearSessionCacheForTests,
  createSession,
  markSessionPaid,
} from '../src/utils/sessionCache.js';

function get(server: Server, path: string): Promise<{
  statusCode: number;
  headers: Record<string, string | string[] | undefined>;
  body: Buffer;
}> {
  const address = server.address();
  assert.ok(address && typeof address === 'object');

  return new Promise((resolve, reject) => {
    const req = request(
      {
        host: '127.0.0.1',
        port: address.port,
        path,
        method: 'GET',
      },
      (res) => {
        const chunks: Buffer[] = [];
        res.on('data', (chunk) => {
          chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
        });
        res.on('end', () => {
          resolve({
            statusCode: res.statusCode ?? 0,
            headers: res.headers,
            body: Buffer.concat(chunks),
          });
        });
      }
    );
    req.on('error', reject);
    req.end();
  });
}

describe('checkout download lock', () => {
  let server: Server;

  before(async () => {
    server = app.listen(0, '127.0.0.1');
    await new Promise<void>((resolve) => server.once('listening', resolve));
  });

  beforeEach(() => {
    process.env.CHECKOUT_REQUIRED = '1';
    clearSessionCacheForTests();
  });

  afterEach(() => {
    delete process.env.CHECKOUT_REQUIRED;
    clearSessionCacheForTests();
  });

  after(async () => {
    await new Promise<void>((resolve, reject) => {
      server.close((err) => (err ? reject(err) : resolve()));
    });
  });

  it('blocks direct file download until the session is paid', async () => {
    const session = createSession(Buffer.from('modified-result'), 'docx', '报价.docx');

    const res = await get(server, `/api/session/${session.id}/download`);

    assert.equal(res.statusCode, 402);
    assert.match(res.body.toString('utf8'), /请先完成付款/);
  });

  it('serves the modified document after checkout marks the session paid', async () => {
    const session = createSession(Buffer.from('modified-result'), 'docx', '报价.docx');
    markSessionPaid(session.id, 'cs_test_paid');

    const res = await get(server, `/api/session/${session.id}/download`);

    assert.equal(res.statusCode, 200);
    assert.match(String(res.headers['content-type']), /wordprocessingml\.document/);
    assert.match(String(res.headers['content-disposition']), /modified_/);
    assert.equal(res.body.toString('utf8'), 'modified-result');
  });
});
