export function resolveFrozenAxisLayout(
  visibleIndices: number[],
  displaySizes: number[],
  freezeBoundary: number,
  headerSize: number,
  viewportSize: number
) {
  const frozenIndices: number[] = [];
  const stickyOffsets = new Map<number, number>();
  let endOffset = headerSize;

  for (const [index, actualIndex] of visibleIndices.entries()) {
    if (actualIndex >= freezeBoundary) break;

    const nextOffset = endOffset + (displaySizes[index] ?? 0);
    // Keep the configured boundary in the sheet, but let the entire axis scroll
    // if freezing it would leave no room for the scrollable pane.
    if (viewportSize > 0 && nextOffset >= viewportSize) {
      return { frozenIndices: [], stickyOffsets: new Map<number, number>(), endOffset: headerSize };
    }

    frozenIndices.push(actualIndex);
    stickyOffsets.set(actualIndex, endOffset);
    endOffset = nextOffset;
  }

  return { frozenIndices, stickyOffsets, endOffset };
}
