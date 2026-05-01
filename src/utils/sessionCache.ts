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
}

const TTL_MS = 10 * 60 * 1000; // 10-minute idle window (sliding)
const MAX_REUSE = 100;

const sessions = new Map<string, DocumentSession>();

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

export interface SessionStatus {
  id: string;
  filename: string;
  fileType: FileType;
  reuseCount: number;
  reuseRemaining: number;
  expiresInMs: number;
  ttlMs: number;
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
  };
}

// Periodic cleanup; .unref() so it doesn't block process exit.
const sweeper = setInterval(() => {
  const now = Date.now();
  let expired = 0;
  for (const [id, s] of sessions) {
    if (now - s.lastUsedAt > TTL_MS) {
      sessions.delete(id);
      expired++;
    }
  }
  if (expired > 0) logger.debug('Session sweeper purged expired', { expired });
}, 60_000);
sweeper.unref();
