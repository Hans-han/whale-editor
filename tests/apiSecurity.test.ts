import { describe, it, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { request } from 'node:http';
import type { Server } from 'node:http';
import { app, isAllowedCorsOrigin } from '../src/api/server.js';

function get(server: Server, path: string, headers: Record<string, string> = {}): Promise<{
  statusCode: number;
  headers: Record<string, string | string[] | undefined>;
  body: string;
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
        headers,
      },
      (res) => {
        let body = '';
        res.setEncoding('utf8');
        res.on('data', (chunk) => {
          body += chunk;
        });
        res.on('end', () => {
          resolve({
            statusCode: res.statusCode ?? 0,
            headers: res.headers,
            body,
          });
        });
      }
    );
    req.on('error', reject);
    req.end();
  });
}

describe('api security headers and CORS', () => {
  let server: Server;

  before(async () => {
    server = app.listen(0, '127.0.0.1');
    await new Promise<void>((resolve) => server.once('listening', resolve));
  });

  after(async () => {
    await new Promise<void>((resolve, reject) => {
      server.close((err) => (err ? reject(err) : resolve()));
    });
  });

  it('sets baseline security headers and hides Express fingerprinting', async () => {
    const res = await get(server, '/api/health');

    assert.equal(res.statusCode, 200);
    assert.equal(res.headers['x-powered-by'], undefined);
    assert.equal(res.headers['x-content-type-options'], 'nosniff');
    assert.equal(res.headers['x-frame-options'], 'SAMEORIGIN');
    assert.equal(res.headers['referrer-policy'], 'no-referrer');
    assert.match(String(res.headers['permissions-policy']), /geolocation=\(\)/);
    assert.match(String(res.headers['content-security-policy']), /default-src 'self'/);
  });

  it('allows local and explicitly configured CORS origins without using wildcard', async () => {
    process.env.CORS_ALLOWED_ORIGINS = 'https://example.test';

    assert.equal(isAllowedCorsOrigin('http://localhost:3000'), true);
    assert.equal(isAllowedCorsOrigin('https://127.0.0.1:3000'), true);
    assert.equal(isAllowedCorsOrigin('https://example.test'), true);
    assert.equal(isAllowedCorsOrigin('https://evil.example'), false);

    const res = await get(server, '/api/health', {
      Origin: 'https://127.0.0.1:3000',
    });
    assert.equal(res.headers['access-control-allow-origin'], 'https://127.0.0.1:3000');
    assert.notEqual(res.headers['access-control-allow-origin'], '*');

    delete process.env.CORS_ALLOWED_ORIGINS;
  });
});
