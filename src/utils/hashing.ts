import { createHash } from 'node:crypto';

export function sha256(data: string): string {
  return createHash('sha256').update(data, 'utf8').digest('hex');
}

export function shortHash(data: string, length = 12): string {
  return sha256(data).slice(0, length);
}
