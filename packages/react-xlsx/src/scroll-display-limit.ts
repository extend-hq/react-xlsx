export function minimumVisibleExtent(minimum: number, sortedHidden: readonly number[] | undefined, maximum: number) {
  if (!sortedHidden?.length || sortedHidden.length === maximum) return minimum;
  let extent = minimum;
  for (const index of sortedHidden) {
    if (index >= extent || extent >= maximum) break;
    extent++;
  }
  return extent;
}

export function growScrollDisplayLimit(
  current: number,
  growth: number,
  maximum: number,
  readOnly: boolean
): number {
  const nextLimit = current + growth;
  return readOnly ? Math.max(current, Math.min(maximum, nextLimit)) : nextLimit;
}
