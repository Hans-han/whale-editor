export function stableStringify(obj: unknown): string {
  return JSON.stringify(obj, stableReplacer);
}

function stableReplacer(key: string, value: unknown): unknown {
  if (value && typeof value === 'object' && !Array.isArray(value)) {
    return Object.keys(value)
      .sort()
      .reduce((sorted: Record<string, unknown>, k) => {
        sorted[k] = (value as Record<string, unknown>)[k];
        return sorted;
      }, {});
  }
  return value;
}

export function stableParse<T>(text: string): T {
  return JSON.parse(text) as T;
}
