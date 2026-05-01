import type { UsageMetrics } from '../types/index.js';

interface DeepSeekUsage {
  prompt_tokens: number;
  completion_tokens: number;
  total_tokens: number;
  prompt_cache_hit_tokens?: number;
  prompt_cache_miss_tokens?: number;
}

export function extractUsage(raw: DeepSeekUsage): UsageMetrics {
  const inputTokens = raw.prompt_tokens ?? 0;
  const outputTokens = raw.completion_tokens ?? 0;
  const totalTokens = raw.total_tokens ?? inputTokens + outputTokens;
  const cacheHitTokens = raw.prompt_cache_hit_tokens;
  const cacheMissTokens = raw.prompt_cache_miss_tokens;

  let cacheHitRatio: number | undefined;
  if (cacheHitTokens !== undefined && cacheMissTokens !== undefined) {
    const total = cacheHitTokens + cacheMissTokens;
    cacheHitRatio = total > 0 ? cacheHitTokens / total : 0;
  }

  return {
    inputTokens,
    outputTokens,
    totalTokens,
    cacheHitTokens,
    cacheMissTokens,
    cacheHitRatio,
  };
}
