import * as React from "react";
import type { XlsxPivotTable, XlsxPivotToggle } from "./pivot-types";
import type { XlsxCellAddress, XlsxImageRect } from "./types";

export function PivotGroupControls({
  pivots, sheetName, rectForCell, viewport, stickyTop, stickyLeft, frozenRight, frozenBottom,
  zoom, readOnly, onToggle, colors, expandLabel, collapseLabel, renderControl
}: {
  pivots: XlsxPivotTable[];
  sheetName?: string;
  rectForCell: (cell: XlsxCellAddress) => XlsxImageRect | null;
  viewport: XlsxImageRect;
  stickyTop: Map<number, number>;
  stickyLeft: Map<number, number>;
  frozenRight: number;
  frozenBottom: number;
  zoom: number;
  readOnly: boolean;
  onToggle: (id: string, toggle: XlsxPivotToggle) => unknown;
  colors: { surface: string; text: string; border: string };
  expandLabel: string;
  collapseLabel: string;
  renderControl?: (cell: XlsxCellAddress, content: React.ReactNode) => React.ReactNode;
}) {
  const [error, setError] = React.useState<string | null>(null);
  React.useEffect(() => { setError(null); }, [pivots, sheetName]);
  if (readOnly) return null;
  const size = Math.max(14, 14 * zoom);
  return <>
    {pivots.filter(pivot => pivot.target.sheet.toLowerCase() === sheetName?.toLowerCase() && !pivot.unsupportedFeatures.length).flatMap(pivot => (
      (pivot.toggles ?? []).flatMap(toggle => {
        const rect = rectForCell(toggle);
        if (!rect) return [];
        const frozenTop = stickyTop.get(toggle.row);
        const frozenLeft = stickyLeft.get(toggle.col);
        const top = !renderControl && frozenTop !== undefined ? frozenTop + viewport.top : rect.top;
        const left = !renderControl && frozenLeft !== undefined ? frozenLeft + viewport.left : rect.left;
        const hasHeader = pivot.headers?.some(header => header.row === toggle.row && header.col === toggle.col);
        const x = left + Math.max(0, rect.width - size - (hasHeader ? size + 6 * zoom : 3 * zoom));
        const y = top + Math.max(0, (rect.height - size) / 2);
        if (!renderControl && (x < viewport.left + (frozenLeft === undefined ? frozenRight : 0)
          || y < viewport.top + (frozenTop === undefined ? frozenBottom : 0)
          || x + size > viewport.left + viewport.width || y + size > viewport.top + viewport.height)) return [];
        if (renderControl && ((frozenLeft === undefined && (x + size < viewport.left - 256 || x > viewport.left + viewport.width + 256))
          || (frozenTop === undefined && (y + size < viewport.top - 256 || y > viewport.top + viewport.height + 256)))) return [];
        const label = `${toggle.collapsed ? expandLabel : collapseLabel}: ${toggle.caption}`;
        const button = <button
          aria-label={label}
          aria-expanded={!toggle.collapsed}
          data-xlsx-pivot-toggle=""
          key={`${pivot.id}:${toggle.row}:${toggle.col}:${toggle.axis}`}
          onClick={event => {
            event.stopPropagation();
            try { onToggle(pivot.id, toggle); setError(null); }
            catch (cause) { setError(cause instanceof Error ? cause.message : String(cause)); }
          }}
          onPointerDown={event => event.stopPropagation()}
          onKeyDown={event => event.stopPropagation()}
          title={label}
          type="button"
          style={{ position: "absolute", left: x, top: y, width: size, height: size, padding: 0,
            display: "grid", placeItems: "center", pointerEvents: "auto", cursor: "pointer",
            background: colors.surface, color: colors.text, border: `1px solid ${colors.border}`,
            borderRadius: 2, font: `${Math.max(11, 11 * zoom)}px/1 system-ui`, zIndex: 25 }}
        >{toggle.collapsed ? "+" : "−"}</button>;
        return [<React.Fragment key={`${pivot.id}:${toggle.row}:${toggle.col}:${toggle.axis}`}>{renderControl ? renderControl(toggle, button) : button}</React.Fragment>];
      })
    ))}
    {error ? <div role="alert" style={{ position: "absolute", left: viewport.left + frozenRight + 8,
      top: viewport.top + frozenBottom + 8, maxWidth: 360, padding: 8, pointerEvents: "auto",
      color: colors.text, background: colors.surface, border: `1px solid ${colors.border}`, zIndex: 30 }}>
      {error}<button type="button" onClick={() => setError(null)} style={{ marginLeft: 8 }}>×</button>
    </div> : null}
  </>;
}
