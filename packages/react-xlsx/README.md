# @extend-ai/react-xlsx

React components and hooks for rendering `.xlsx` workbooks in the browser.

`@extend-ai/react-xlsx` gives you:

- A drop-in `XlsxViewer` for workbook previews
- A provider/controller API for custom spreadsheet experiences
- Worksheet rendering with frozen panes, tables, merged cells, conditional formatting, sparklines, selection, resizing, copy/paste, and zoom
- Embedded worksheet images, shapes, form controls, charts, and chartsheet tabs
- Worker-backed parsing and large-file guardrails
- Thumbnail helpers for building sheet strips, previews, and navigation UIs
- TypeScript types for viewer state, workbook metadata, charts, images, tables, and render hooks

## Install

```bash
npm install @extend-ai/react-xlsx react react-dom
```

```bash
pnpm add @extend-ai/react-xlsx react react-dom
```

`react` and `react-dom` are peer dependencies.

## WebAssembly Asset

Workbook parsing and calculation run through the `@dukelib/sheets-wasm` WebAssembly module. The module loads lazily, on the first workbook parse.

Most apps can use the default loader. If your bundler or deployment needs to host the WASM binary somewhere explicit, configure it before the first parse:

```ts
import { setWasmSource } from "@extend-ai/react-xlsx";

setWasmSource("https://cdn.example.com/duke_sheets_wasm_bg.wasm");
// or pass a URL, Request, Response, ArrayBuffer/TypedArray, or compiled WebAssembly.Module
```

The Duke WASM binary is also exposed as a package subpath:

```ts
import wasmUrl from "@extend-ai/react-xlsx/duke_sheets_wasm_bg.wasm?url";
import { setWasmSource } from "@extend-ai/react-xlsx";

setWasmSource(wasmUrl);
```

### Next.js Turbopack

Turbopack may try to treat `.wasm?url` imports as WebAssembly modules during static analysis. For Turbopack apps, use a plain public or CDN URL instead of importing the WASM file with `?url`.

Copy the WASM file into your app's `public/` directory:

```bash
cp node_modules/@extend-ai/react-xlsx/dist/duke_sheets_wasm_bg.wasm public/duke_sheets_wasm_bg.wasm
```

Then configure the source from a shared client module before any workbook is parsed:

```ts
// app/xlsx-wasm.ts
"use client";

import { setWasmSource } from "@extend-ai/react-xlsx";

setWasmSource("/duke_sheets_wasm_bg.wasm");
```

Import that setup module before rendering any XLSX viewer, provider, or controller:

```tsx
// app/workbook-preview.tsx
"use client";

import "./xlsx-wasm";
import { XlsxViewer } from "@extend-ai/react-xlsx";

export function WorkbookPreview({ file }: { file: ArrayBuffer }) {
  return <XlsxViewer file={file} height={600} />;
}
```

If several routes use the viewer, import the same setup module from a shared client boundary such as `app/providers.tsx`. Calling `setWasmSource()` more than once with the same source is fine before initialization, but the source must not change after the first parse because the WASM module is initialized once per JavaScript context.

Configured string, URL, Request URL, bytes, and `WebAssembly.Module` sources are forwarded into the XLSX worker. `Response` sources are supported on the main thread; worker-backed parsing is skipped for that source type.

You can also call `initWasm()` (optionally with a source) ahead of time to warm the module before the first workbook is opened.

## Main Entry Points

The package exports three useful levels of API:

1. `XlsxViewer`
   A ready-to-render workbook viewer with built-in toolbar, sheet tabs, grid rendering, charts, images, and selection state.

2. `XlsxViewerProvider` + viewer hooks
   Shared controller context for custom toolbars, side panels, thumbnail strips, or other UI around the workbook.

3. `useXlsxViewerController`
   A lower-level controller hook for fully controlled integrations.

## Quick Start

### Basic viewer

Use `XlsxViewer` when you want the smallest integration surface.

```tsx
import * as React from "react";
import { XlsxViewer } from "@extend-ai/react-xlsx";

export function WorkbookPreview() {
  const [file, setFile] = React.useState<ArrayBuffer | undefined>();
  const [fileName, setFileName] = React.useState<string | undefined>();

  return (
    <div style={{ display: "grid", gap: 12 }}>
      <input
        type="file"
        accept=".xlsx,.xlsm,.xls"
        onChange={async (event) => {
          const nextFile = event.target.files?.[0];
          setFile(nextFile ? await nextFile.arrayBuffer() : undefined);
          setFileName(nextFile?.name);
        }}
      />

      <XlsxViewer
        file={file}
        fileName={fileName}
        height={600}
        emptyState="Choose a workbook to preview."
      />
    </div>
  );
}
```

You can also load a remote workbook with `src`:

```tsx
import { XlsxViewer } from "@extend-ai/react-xlsx";

export function RemoteWorkbookPreview() {
  return <XlsxViewer src="/reports/quarterly-model.xlsx" height="70vh" />;
}
```

### Provider and hooks

Use `XlsxViewerProvider` when custom UI needs access to the active workbook, selection, zoom, charts, images, tables, or editing commands.

```tsx
import {
  DefaultXlsxToolbar,
  XlsxViewer,
  XlsxViewerProvider,
  useXlsxViewerSelection,
  useXlsxViewerZoom,
} from "@extend-ai/react-xlsx";

function WorkbookStatus() {
  const { activeCellAddress, selectedRangeAddress } = useXlsxViewerSelection();
  const { zoomScale, zoomIn, zoomOut, canZoomIn, canZoomOut } = useXlsxViewerZoom();

  return (
    <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
      <span>{selectedRangeAddress ?? activeCellAddress ?? "No selection"}</span>
      <button type="button" onClick={zoomOut} disabled={!canZoomOut}>
        -
      </button>
      <span>{zoomScale}%</span>
      <button type="button" onClick={zoomIn} disabled={!canZoomIn}>
        +
      </button>
    </div>
  );
}

export function WorkbookWorkspace({ file }: { file: ArrayBuffer }) {
  return (
    <XlsxViewerProvider file={file} fileName="model.xlsx">
      <DefaultXlsxToolbar />
      <WorkbookStatus />
      <XlsxViewer height="70vh" showDefaultToolbar={false} />
    </XlsxViewerProvider>
  );
}
```

### Controlled controller

Use `useXlsxViewerController` when you want to own the controller instance and pass it into several components.

```tsx
import {
  XlsxViewer,
  useXlsxViewerController,
} from "@extend-ai/react-xlsx";

export function ControlledWorkbook({ file }: { file: ArrayBuffer }) {
  const controller = useXlsxViewerController({
    file,
    fileName: "forecast.xlsx",
    readOnlyAboveBytes: 10 * 1024 * 1024,
  });

  return (
    <div style={{ display: "grid", gap: 12 }}>
      <button type="button" onClick={controller.exportXlsx} disabled={!controller.canExport}>
        Export XLSX
      </button>

      <XlsxViewer controller={controller} height={640} />
    </div>
  );
}
```

## Useful Hooks

These hooks work inside `XlsxViewer` or `XlsxViewerProvider` context.

- `useXlsxViewer()` for full controller access
- `useXlsxViewerSelection()` for active cell and range state
- `useXlsxViewerZoom()` for zoom controls and limits
- `useXlsxViewerEditing()` for editing, persisted style writes, formulas, undo/redo, fill, merge, clipboard, and export actions
- `useXlsxViewerTables()` for table metadata and table sorting
- `useXlsxViewerImages()` for embedded image and chart selection, movement, and resizing
- `useXlsxViewerCharts()` for chart, chart element, formula, and chartsheet state
- `useXlsxViewerThumbnails(options)` for painting worksheet thumbnails into your own canvases

## Chart Selection And Formulas

Chart interactions follow Excel's selection progression: clicking a chart selects the chart frame, clicking a plotted item or legend entry again selects the series, and clicking an item in an already-selected series selects that point. `selectedChartElement` exposes the current chart, series, legend entry, or point selection through the controller and `useXlsxViewerCharts()`.

When a series, legend entry, or point is selected, `selectedFormula` returns an Excel-style `=SERIES(...)` formula and `selectedFormulaTarget` is `{ kind: "chartSeries", chartId, seriesId, seriesIndex }`. Use `setSelectedFormula(...)` for a formula bar that edits either the active cell formula or the selected chart series formula. Use `getChartSeriesFormula(chartId, seriesIndex)` and `setChartSeriesFormula(chartId, seriesIndex, formula)` when you need direct chart-series formula control.

## Thumbnail Hook

`useXlsxViewerThumbnails` returns paint functions for each worksheet so you can build your own sheet strip or navigation UI.

```tsx
import * as React from "react";
import {
  XlsxViewerProvider,
  useXlsxViewerThumbnails,
  type XlsxSheetThumbnail,
} from "@extend-ai/react-xlsx";

function SheetThumbnail({ thumbnail }: { thumbnail: XlsxSheetThumbnail }) {
  const canvasRef = React.useRef<HTMLCanvasElement | null>(null);

  React.useEffect(() => {
    thumbnail.paint(canvasRef.current);
  }, [thumbnail]);

  return (
    <canvas
      ref={canvasRef}
      width={thumbnail.width}
      height={thumbnail.height}
      style={{ width: thumbnail.width, height: thumbnail.height }}
    />
  );
}

function SheetThumbnailStrip() {
  const { thumbnails } = useXlsxViewerThumbnails({
    includeHeaders: true,
    resolution: { maxWidth: 180, maxHeight: 120 },
  });

  return (
    <div style={{ display: "flex", gap: 12, overflowX: "auto" }}>
      {thumbnails.map((thumbnail) => (
        <SheetThumbnail
          key={thumbnail.workbookSheetIndex}
          thumbnail={thumbnail}
        />
      ))}
    </div>
  );
}

export function ThumbnailExample({ file }: { file: ArrayBuffer }) {
  return (
    <XlsxViewerProvider file={file}>
      <SheetThumbnailStrip />
    </XlsxViewerProvider>
  );
}
```

Notes:

- `resolution` accepts either a single max dimension or `{ maxWidth, maxHeight }`.
- Thumbnails preserve worksheet aspect ratio and paint into your supplied `<canvas>`.
- The current implementation renders a bounded top-left worksheet preview, including loaded embedded worksheet images, shapes, and form controls, but does not include charts.

## Custom Rendering

The viewer exposes render props for common UI integration points.

```tsx
import { XlsxViewer } from "@extend-ai/react-xlsx";

export function CustomWorkbook({ file }: { file: ArrayBuffer }) {
  return (
    <XlsxViewer
      file={file}
      height={600}
      selectionColor="#2563eb"
      renderImage={({ image, style }) => (
        <img
          src={image.src}
          alt={image.description ?? image.name ?? ""}
          style={{ ...style, objectFit: "contain" }}
        />
      )}
      renderTableHeaderMenu={({ column, direction, sortAscending, sortDescending, triggerIcon, triggerProps }) => (
        <span>
          <button type="button" {...triggerProps}>
            {triggerIcon}
          </button>
          <button type="button" onClick={sortAscending}>
            Sort A to Z{direction === "ascending" ? " selected" : ""}
          </button>
          <button type="button" onClick={sortDescending}>
            Sort Z to A{direction === "descending" ? " selected" : ""}
          </button>
          <span>{column.name}</span>
        </span>
      )}
    />
  );
}
```

Apply `triggerProps` to the table-header trigger button so clicks do not leak into grid selection.

## Large Files

`XlsxViewer` includes guardrails for large workbooks.

```tsx
import { XlsxViewer } from "@extend-ai/react-xlsx";

export function LargeWorkbookPreview({ file }: { file: ArrayBuffer }) {
  return (
    <XlsxViewer
      file={file}
      maxFileSizeBytes={50 * 1024 * 1024}
      readOnlyAboveBytes={10 * 1024 * 1024}
      deferLoadingAboveBytes={20 * 1024 * 1024}
      fileTooLargeState={({ displayFileName, fileSizeBytes, maxFileSizeBytes }) => (
        <div>
          <strong>{displayFileName}</strong> is too large to open here.
          <div>
            {Math.round(fileSizeBytes / (1024 * 1024))} MB of{" "}
            {Math.round(maxFileSizeBytes / (1024 * 1024))} MB allowed
          </div>
        </div>
      )}
    />
  );
}
```

Notes:

- `maxFileSizeBytes` defaults to `25 MB`.
- `readOnlyAboveBytes` can disable mutation actions for larger files.
- `deferLoadingAboveBytes` waits for `controller.continueDeferredLoad()` before parsing files above the threshold.
- `useWorker` defaults to `true` when browser workers are available.

## Viewer Props

`XlsxViewerProps` includes all controller options plus rendering options.

Common source and loading props:

- `file?: ArrayBuffer`
- `src?: string`
- `fileName?: string`
- `controller?: XlsxViewerController`
- `useWorker?: boolean`
- `maxFileSizeBytes?: number`
- `readOnly?: boolean`
- `readOnlyAboveBytes?: number`
- `deferLoadingAboveBytes?: number`
- `showHiddenSheets?: boolean`
- `skipXmlParsing?: boolean`

Common rendering props:

- `height?: React.CSSProperties["height"]`
- `className?: string`
- `isDark?: boolean`
- `rounded?: boolean`
- `showDefaultToolbar?: boolean`
- `toolbar?: React.ReactNode | ((controller: XlsxViewerController) => React.ReactNode)`
- `experimentalCanvas?: boolean`
- `enableGestureZoom?: boolean`
- `enableCanvasSelectionAnimation?: boolean`
- `allowResizeInReadOnly?: boolean`
- `selectionColor?: string`
- `selectionFillColor?: string`
- `selectionHeaderColor?: string`
- `showImages?: boolean`
- `emptyState?: React.ReactNode`
- `loadingState?: React.ReactNode`
- `errorState?: React.ReactNode | ((error: Error) => React.ReactNode)`
- `fileTooLargeState?: React.ReactNode | ((props: XlsxFileTooLargeRenderProps) => React.ReactNode)`
- `getCellStyle?: (context: XlsxCellStyleContext) => React.CSSProperties | null | undefined`
- `renderImage?: (props: XlsxImageRenderProps) => React.ReactNode`
- `renderImageSelection?: (props: XlsxImageSelectionRenderProps) => React.ReactNode`
- `renderChartLoading?: (props: XlsxChartLoadingRenderProps) => React.ReactNode`
- `renderTableHeaderMenu?: (props: XlsxTableHeaderMenuRenderProps) => React.ReactNode`
- `renderScroller?: (props: XlsxScrollerRenderProps) => React.ReactNode`

### Persisted Cell Styling

Use `setCellStyle`, `setSelectedCellStyle`, and `setRangeStyle` when a custom toolbar should write Excel formatting into the workbook. These APIs mutate workbook data, participate in undo/redo, refresh the viewer, and are included in `exportXlsx()`.

```tsx
import {
  useXlsxViewer,
  type XlsxCellStyleInput
} from "@extend-ai/react-xlsx";

const highlightStyle: XlsxCellStyleInput = {
  font: { bold: true, color: { colorType: "rgb", hex: "1D4ED8" } },
  fill: { fillType: "solid", color: { colorType: "rgb", hex: "DBEAFE" } },
  alignment: { horizontal: "center", vertical: "center", wrapText: true }
};

function FormattingButton() {
  const { selection, setRangeStyle, setSelectedCellStyle } = useXlsxViewer();

  return (
    <button
      type="button"
      onClick={() => {
        if (selection) {
          setRangeStyle(selection, highlightStyle);
          return;
        }
        setSelectedCellStyle(highlightStyle);
      }}
    >
      Highlight
    </button>
  );
}
```

`XlsxCellStyleInput` supports these persisted Excel style groups:

| Field | Type | Notes |
| --- | --- | --- |
| `font` | `XlsxCellFontStyleInput` | Font family, size, bold, italic, underline, strikethrough, color, superscript/subscript. |
| `fill` | `XlsxCellFillStyleInput` | Solid, pattern, and gradient fills. |
| `border` | `XlsxCellBorderStyleInput` | Per-edge borders, colors, styles, and diagonal borders. |
| `alignment` | `XlsxCellAlignmentInput` | Horizontal/vertical alignment, wrap text, shrink to fit, indent, rotation, reading order. |
| `numberFormat` | `XlsxCellNumberFormatInput` | General, builtin, or custom Excel number format strings. |
| `protection` | `XlsxCellProtectionInput` | Locked/hidden flags used when sheet protection is enabled. |

### Render-Only Cell Styling

`getCellStyle` is an escape hatch for styling individual cells without forking the workbook data. It is called for every rendered cell and returns a partial `React.CSSProperties` that merges on top of the viewer's resolved style. Return `undefined` (or `null`) to leave a cell untouched.

```tsx
import { XlsxViewer, type XlsxViewerProps } from "@extend-ai/react-xlsx";

function Workbook({ buffer, highlighted }: { buffer: ArrayBuffer; highlighted: Set<string> }) {
  const getCellStyle = React.useCallback<NonNullable<XlsxViewerProps["getCellStyle"]>>(
    ({ cell, isTableHeader }) => {
      if (isTableHeader) {
        return undefined;
      }
      if (highlighted.has(`${cell.row}:${cell.col}`)) {
        return { backgroundColor: "rgba(37, 99, 235, 0.12)", outline: "1px solid #2563eb" };
      }
      return undefined;
    },
    [highlighted]
  );

  return <XlsxViewer file={buffer} getCellStyle={getCellStyle} />;
}
```

The `context` passed to `getCellStyle` is an `XlsxCellStyleContext`:

| Field | Type | Notes |
| --- | --- | --- |
| `cell` | `XlsxCellAddress` | Address of the cell being styled. |
| `workbookSheetIndex` | `number` | Workbook sheet index of the cell's sheet. |
| `sheetName` | `string` | Display name of the cell's sheet. |
| `resolvedStyle` | `React.CSSProperties` | The style the viewer computed (workbook formatting + built-ins). Read-only. |
| `value` | `string` | The cell's resolved display value. |
| `hasValidation` | `boolean` | Cell has a data validation rule. |
| `hasHyperlink` | `boolean` | Cell has a hyperlink. |
| `hasConditionalFormat` | `boolean` | Cell is affected by a color scale, data bar, or icon set. |
| `hasChartHighlight` | `boolean` | Cell is in a selected chart's highlighted source range. |
| `isMerged` | `boolean` | Cell is the anchor of a merged range. |
| `isTableHeader` | `boolean` | Cell is a table header cell. |

Notes:

- Keep the callback stable (e.g. `useCallback`) so cell styling is not recomputed every render. When the callback identity changes, the viewer re-resolves and repaints cells.
- The DOM renderer (`experimentalCanvas={false}`) applies every returned CSS property.
- The canvas renderer (the default) honors the subset it can paint: `backgroundColor`, `backgroundImage` gradients, `color`, the four `border*` sides, `padding`, `textAlign`, `textDecoration`, `textOverflow`, and font properties. CSS-only effects such as `boxShadow`, `outline`, or `animation` apply in the DOM renderer.
- `getCellStyle` is not applied to worksheet thumbnails painted via `useXlsxViewerThumbnails(...)`.

### Custom Scroll Area

By default, the viewer renders its native scroll viewport with the browser scrollbar. To use a custom scroll area, provide `renderScroller` and spread `viewportProps` onto the actual scrollable viewport element:

```tsx
import { ScrollArea as ScrollAreaPrimitive } from "@base-ui/react/scroll-area";
import { XlsxViewer } from "@extend-ai/react-xlsx";

function Workbook() {
  return (
    <XlsxViewer
      src="/model.xlsx"
      renderScroller={({ children, viewportProps }) => (
        <ScrollAreaPrimitive.Root className="h-full min-h-0 w-full min-w-0 flex-1">
          <ScrollAreaPrimitive.Viewport {...viewportProps}>
            {children}
          </ScrollAreaPrimitive.Viewport>
          <ScrollAreaPrimitive.Scrollbar orientation="vertical">
            <ScrollAreaPrimitive.Thumb />
          </ScrollAreaPrimitive.Scrollbar>
          <ScrollAreaPrimitive.Scrollbar orientation="horizontal">
            <ScrollAreaPrimitive.Thumb />
          </ScrollAreaPrimitive.Scrollbar>
          <ScrollAreaPrimitive.Corner />
        </ScrollAreaPrimitive.Root>
      )}
    />
  );
}
```

`viewportProps` includes the viewer ref, scroll handler, keyboard/copy/paste handlers, focus state, and required sizing styles. Applying those props to the scroll area root instead of the viewport can break virtualization, canvas synchronization, and keyboard navigation.

## Workbook Support

Primary support is for OOXML `.xlsx` workbooks.

Supported worksheet features include:

- Frozen panes, merged cells, row and column sizing, hidden rows and columns, and gridline settings
- Cell styles, fills, borders, alignment, number formats, formulas, cached formula values, and cell controls
- Tables, table sorting, conditional formatting, data validation metadata, and sparklines
- Embedded images, shapes, form controls, worksheet charts, and chartsheet tabs
- Copy, paste, undo, redo, merge, unmerge, fill, CSV export, and XLSX export

Chart rendering supports common Excel chart families including column, bar, line, area, scatter, pie, doughnut, radar, bubble, stock, surface, waterfall, funnel, box-and-whisker, sunburst, treemap, region map, combo charts, and chartsheets. Selecting a chart series highlights its workbook source ranges; point selection keeps the series ranges highlighted and emphasizes the matching point cells when the source references resolve to rectangular ranges.

Legacy `.xls` and macro-enabled `.xlsm` files have limited support. The viewer only displays workbook data that `@dukelib/sheets-wasm` can parse, and format-specific XML features may be missing or skipped.

## Exported Types

The package exports the main types you are likely to use for custom integrations:

- `UseXlsxViewerControllerOptions`
- `XlsxViewerProps`
- `XlsxViewerProviderProps`
- `XlsxViewerController`
- `XlsxViewerSelection`
- `XlsxViewerZoom`
- `XlsxViewerEditing`
- `XlsxViewerTables`
- `XlsxViewerImages`
- `XlsxViewerCharts`
- `XlsxViewerThumbnails`
- `XlsxScrollerRenderProps`
- `XlsxCellStyleContext`
- `XlsxCellStyleInput`, `XlsxCellFontStyleInput`, `XlsxCellFillStyleInput`, `XlsxCellBorderStyleInput`
- `XlsxSheetThumbnail`
- `UseXlsxViewerThumbnailsOptions`
- `XlsxChart`, `XlsxChartSeries`, `XlsxChartAxis`, `XlsxChartElementSelection`, `XlsxChartsheet`
- `XlsxFormulaTarget`
- `XlsxImage`, `XlsxImageRect`, `XlsxImageRenderProps`, `XlsxImageSelectionRenderProps`
- `XlsxTable`, `XlsxTableColumn`, `XlsxTableHeaderMenuRenderProps`
- `XlsxWorkbookTab`, `XlsxCellAddress`, `XlsxCellRange`

## License

See the repository license for usage terms.
