export function growScrollDisplayLimit(
  current: number,
  growth: number,
  maximum: number,
  readOnly: boolean
): number {
  const nextLimit = current + growth;
  return readOnly ? Math.max(current, Math.min(maximum, nextLimit)) : nextLimit;
}
