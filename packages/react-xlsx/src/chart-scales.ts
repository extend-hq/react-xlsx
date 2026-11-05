type Span = readonly [number, number];

export function linearScale([low, high]: Span, [start, end]: Span): (value: number) => number {
  if (high < low) return linearScale([high, low], [end, start]);
  const span = high - low;
  return (value: number): number => {
    const t = span === 0 ? 0.5 : (value - low) / span;
    return start * (1 - t) + end * t;
  };
}

export function bandScale<T extends string | number>(
  domain: readonly T[], [from, to]: Span, inner: number, outer: number
) {
  const indices = new Map<T, number>();
  for (const value of domain) {
    if (!indices.has(value)) indices.set(value, indices.size);
  }
  const count = indices.size;
  const width = Math.abs(to - from);
  const step = width / Math.max(1, count - inner + 2 * outer);
  const origin = Math.min(from, to) + (width - step * (count - inner)) / 2;
  const locate = (value: T) => {
    const index = indices.get(value);
    return index == null ? undefined : origin + step * (to < from ? count - 1 - index : index);
  };
  return Object.assign(locate, { bandwidth: step * (1 - inner) });
}

export function pointScale<T extends string | number>(domain: readonly T[], range: Span) {
  return bandScale(domain, range, 1, 0);
}
