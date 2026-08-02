export interface CellTextClipOverscan {
  horizontal: number;
  vertical: number;
}

export function resolveCellTextClipOverscan(
  overscan: number,
  usesWrappedText: boolean
): CellTextClipOverscan {
  return {
    horizontal: overscan,
    vertical: usesWrappedText ? 0 : overscan
  };
}
