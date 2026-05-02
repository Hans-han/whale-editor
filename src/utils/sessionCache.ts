import { randomUUID } from 'node:crypto';
import type { FileType } from '../types/index.js';
import { logger } from './logger.js';

export interface DocumentSession {
  id: string;
  buffer: Buffer;
  fileType: FileType;
  filename: string;
  createdAt: number;
  lastUsedAt: number;
  reuseCount: number;
  paidAt?: number;
  checkoutSessionId?: string;
}

const TTL_MS = 10 * 60 * 1000; // 10-minute idle window (sliding)
const MAX_REUSE = 100;
const DEFAULT_MAX_SESSIONS = 50;
const DEFAULT_MAX_TOTAL_BYTES = 250 * 1024 * 1024;

const sessions = new Map<string, DocumentSession>();

function parsePositiveInt(value: string | undefined, fallback: number): number {
  const parsed = Number.parseInt(value ?? '', 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

function maxSessions(): number {
  return parsePositiveInt(process.env.SESSION_CACHE_MAX_SESSIONS, DEFAULT_MAX_SESSIONS);
}

function maxTotalBytes(): number {
  return parsePositiveInt(process.env.SESSION_CACHE_MAX_BYTES, DEFAULT_MAX_TOTAL_BYTES);
}

function totalBytes(): number {
  let total = 0;
  for (const session of sessions.values()) {
    total += session.buffer.length;
  }
  return total;
}

function purgeExpired(now = Date.now()): number {
  let expired = 0;
  for (const [id, s] of sessions) {
    if (now - s.lastUsedAt > TTL_MS) {
      sessions.delete(id);
      expired++;
    }
  }
  return expired;
}

function oldestSessionId(exceptId?: string): string | null {
  let oldestId: string | null = null;
  let oldestLastUsed = Number.POSITIVE_INFINITY;
  let oldestCreated = Number.POSITIVE_INFINITY;

  for (const [id, session] of sessions) {
    if (id === exceptId) continue;
    if (
      session.lastUsedAt < oldestLastUsed ||
      (session.lastUsedAt === oldestLastUsed && session.createdAt < oldestCreated)
    ) {
      oldestId = id;
      oldestLastUsed = session.lastUsedAt;
      oldestCreated = session.createdAt;
    }
  }

  return oldestId;
}

function enforceLimits(protectedId?: string): void {
  const expired = purgeExpired();
  if (expired > 0) logger.debug('Session cache purged expired before limit check', { expired });

  const countLimit = maxSessions();
  const byteLimit = maxTotalBytes();
  let evicted = 0;

  while (sessions.size > countLimit || totalBytes() > byteLimit) {
    const id = oldestSessionId(protectedId) ?? oldestSessionId();
    if (!id) break;
    sessions.delete(id);
    evicted++;
  }

  if (evicted > 0) {
    logger.info('Session cache evicted oldest sessions', {
      evicted,
      sessions: sessions.size,
      totalBytes: totalBytes(),
      maxSessions: countLimit,
      maxTotalBytes: byteLimit,
    });
  }
}

export function createSession(
  buffer: Buffer,
  fileType: FileType,
  filename: string
): DocumentSession {
  const id = randomUUID();
  const now = Date.now();
  const session: DocumentSession = {
    id,
    buffer,
    fileType,
    filename,
    createdAt: now,
    lastUsedAt: now,
    reuseCount: 0,
  };
  sessions.set(id, session);
  enforceLimits(id);
  logger.info('Session created', { id, fileType, filename, size: buffer.length });
  return session;
}

export function getSession(id: string): DocumentSession | null {
  const s = sessions.get(id);
  if (!s) return null;
  if (Date.now() - s.lastUsedAt > TTL_MS) {
    sessions.delete(id);
    logger.info('Session expired on access', { id, age: Date.now() - s.lastUsedAt });
    return null;
  }
  if (s.reuseCount >= MAX_REUSE) {
    logger.info('Session reuse cap reached', { id, reuseCount: s.reuseCount });
    return null;
  }
  return s;
}

export function bumpSession(id: string): DocumentSession | null {
  const s = sessions.get(id);
  if (!s) return null;
  s.lastUsedAt = Date.now();
  s.reuseCount += 1;
  return s;
}

export function replaceSessionDocument(
  id: string,
  buffer: Buffer,
  fileType: FileType,
  filename: string
): DocumentSession | null {
  const s = sessions.get(id);
  if (!s) return null;
  s.buffer = buffer;
  s.fileType = fileType;
  s.filename = filename;
  s.lastUsedAt = Date.now();
  enforceLimits(id);
  return s;
}

export function markSessionCheckoutStarted(id: string, checkoutSessionId: string): DocumentSession | null {
  const s = sessions.get(id);
  if (!s) return null;
  s.checkoutSessionId = checkoutSessionId;
  s.lastUsedAt = Date.now();
  return s;
}

export function markSessionPaid(id: string, checkoutSessionId?: string): DocumentSession | null {
  const s = sessions.get(id);
  if (!s) return null;
  s.paidAt = Date.now();
  if (checkoutSessionId) s.checkoutSessionId = checkoutSessionId;
  s.lastUsedAt = Date.now();
  return s;
}

export interface SessionStatus {
  id: string;
  filename: string;
  fileType: FileType;
  reuseCount: number;
  reuseRemaining: number;
  expiresInMs: number;
  ttlMs: number;
  paid: boolean;
}

export function statusOf(s: DocumentSession): SessionStatus {
  return {
    id: s.id,
    filename: s.filename,
    fileType: s.fileType,
    reuseCount: s.reuseCount,
    reuseRemaining: Math.max(0, MAX_REUSE - s.reuseCount),
    expiresInMs: Math.max(0, TTL_MS - (Date.now() - s.lastUsedAt)),
    ttlMs: TTL_MS,
    paid: Boolean(s.paidAt),
  };
}

// Periodic cleanup; .unref() so it doesn't block process exit.
const sweeper = setInterval(() => {
  const expired = purgeExpired();
  if (expired > 0) logger.debug('Session sweeper purged expired', { expired });
}, 60_000);
sweeper.unref();

export function sessionCacheStats(): {
  count: number;
  totalBytes: number;
  maxSessions: number;
  maxTotalBytes: number;
} {
  return {
    count: sessions.size,
    totalBytes: totalBytes(),
    maxSessions: maxSessions(),
    maxTotalBytes: maxTotalBytes(),
  };
}

export function clearSessionCacheForTests(): void {
  sessions.clear();
}
