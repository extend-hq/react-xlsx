# @extend-ai/react-xlsx

React components, hooks, and a vanilla JavaScript API for rendering `.xlsx` workbooks in the browser.

`@extend-ai/react-xlsx` gives you:

- A drop-in `XlsxViewer` for workbook previews
- A `createXlsxViewer` API for mounting the viewer from plain JavaScript
- A provider/controller API for custom spreadsheet experiences
- Worksheet rendering with frozen panes, tables, merged cells, conditional formatting, sparklines, selection, resizing, copy/paste, and zoom
- Embedded worksheet images, shapes, form controls, charts, and chartsheet tabs
- Worker-backed parsing and large-file guardrails
- Thumbnail helpers for building sheet strips, previews, and navigation UIs
- TypeScript types for viewer state, workbook metadata, charts, images, tables, and render hooks

## Install

```bash
npm install @extend-ai/react-xlsx
```

```bash
pnpm add @extend-ai/react-xlsx
```

For the React entry, install `react` and `react-dom` as well. They are optional peer dependencies and are not needed by the vanilla entry.

## Automatic Chart Loading

Chart rendering loads automatically when the viewer encounters charts in the active tab. Map geometry and WebGL surface rendering load separately when those chart types are needed. No extra imports or configuration are required.

The viewer starts loading chart modules while chart metadata is being hydrated. The grid stays usable while charts load. Use `renderChartLoading` to supply a custom component for chart metadata, chart rendering code, and map geometry loading. Its `stage` field is `"metadata"`, `"renderer"`, or `"map"`, and its `chart`, `rect`, and `defaultNode` fields provide context. The component renders in the full chart rectangle and can contain ordinary HTML.

```tsx
<XlsxViewer renderChartLoading={() => <LoadingSpinner />} />
```

Surface charts use their SVG rendering while the WebGL module loads. Loaded modules are reused across tabs and workbooks.

Use a production bundler that preserves dynamic imports, and deploy all generated JavaScript chunks alongside the application. This reduces the initial browser download; all chart modules remain included in the installed package.

## Spreadsheet Fonts

The first workbook load also loads a bundled Carlito font fallback before measuring text and column widths. It supplies Calibri-compatible metrics when Calibri is unavailable, including regular, bold, italic, and bold italic styles. Installed Calibri and Carlito fonts remain preferred. Other font families still depend on the host application or system.

The fallback is a separate JavaScript chunk with embedded WOFF2 data, so production rendering does not contact a font CDN. Deploy this chunk with the package's other generated chunks. Carlito is distributed under the SIL Open Font License; the published package includes `dist/Carlito-OFL.txt`.

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

In Vite development, exclude the vanilla entry from dependency optimization so Vite can resolve the worker asset. Production builds do not need this setting:

```ts
import { defineConfig } from "vite";

export default defineConfig({
  optimizeDeps: { exclude: ["@extend-ai/react-xlsx/vanilla"] }
});
```

For the React entry, exclude the shared entries and let Vite optimize React separately:

```ts
import { defineConfig } from "vite";

export default defineConfig({
  optimizeDeps: {
    exclude: ["@extend-ai/react-xlsx", "@extend-ai/react-xlsx/vanilla", "@extend-ai/react-xlsx/internal"],
    include: ["react", "react-dom", "react-dom/client", "react/jsx-runtime"]
  }
});
```

You can also call `initWasm()` (optionally with a source) ahead of time to warm the module before the first workbook is opened. Vanilla applications can import these functions from `@extend-ai/react-xlsx/vanilla`.

## Main Entry Points

The package provides a vanilla JavaScript entry point and three levels of React API:

- `@extend-ai/react-xlsx/vanilla` exports `createXlsxViewer`, the `XlsxViewer` instance class, and utilities for JavaScript applications.

`@extend-ai/react-xlsx` exports:

1. `XlsxViewer`
   A ready-to-render workbook viewer with built-in toolbar, sheet tabs, grid rendering, charts, images, and selection state.

2. `XlsxViewerProvider` + viewer hooks
   Shared controller context for custom toolbars, side panels, thumbnail strips, or other UI around the workbook.

3. `useXlsxViewerController`
   A lower-level controller hook for fully controlled integrations.

## Quick Start

### Vanilla JavaScript

Use the vanilla entry point to mount the viewer in an existing DOM element. The viewer uses a bundled Preact renderer and does not require React or JSX. The React component wraps the same viewer implementation. The instance creates a `div` in the regular DOM, so application styles and focus behavior continue to apply.

The examples below use a browser bundler such as Vite to resolve the package import. For Vite development, use the [worker configuration above](#webassembly-asset).

#### Open a workbook from a URL

```html
<div id="viewer"></div>
<script type="module" src="/main.js"></script>
```

```js
import { createXlsxViewer } from "@extend-ai/react-xlsx/vanilla";

const viewer = createXlsxViewer(document.getElementById("viewer"), {
  src: "/workbook.xlsx",
  height: 600,
  readOnly: true,
});
```

The URL must be accessible to the browser. A workbook on another origin needs the appropriate CORS headers. Omit `readOnly` to allow editing.

You can also construct an instance directly:

```js
import { XlsxViewer } from "@extend-ai/react-xlsx/vanilla";

const viewer = new XlsxViewer(document.getElementById("viewer"), {
  src: "/workbook.xlsx",
  height: 600,
});
```

#### Open a local workbook

The `file` option accepts workbook bytes as an `ArrayBuffer`. This standalone example starts with an empty viewer and loads the file selected by the user:

```html
<input id="workbook" type="file" accept=".xlsx,.xlsm,.xls" />
<div id="viewer"></div>
<script type="module" src="/main.js"></script>
```

```js
import { createXlsxViewer } from "@extend-ai/react-xlsx/vanilla";

const viewer = createXlsxViewer(document.getElementById("viewer"), {
  height: 600,
  emptyState: "Choose a workbook to preview.",
});

document.getElementById("workbook").addEventListener("change", async (event) => {
  const file = event.currentTarget.files?.[0];
  if (!file) return;

  viewer.setOptions({ file: await file.arrayBuffer(), fileName: file.name });
});
```

#### Read state and control the viewer

Subscribe to loading, errors, sheet changes, selection, and other controller updates:

```js
const unsubscribe = viewer.subscribe((controller) => {
  if (controller.error) {
    console.error(controller.error);
  } else if (controller.isLoading) {
    console.log("Loading workbook…");
  } else if (controller.sheets.length > 0) {
    console.log(controller.activeSheet?.name, controller.selectedRangeAddress);
  }
});
```

Read the latest controller when invoking a command. These commands should run after the workbook has loaded:

```js
const controller = viewer.getController();

if (controller && !controller.isLoading && controller.sheets.length > 0) {
  controller.highlightRanges("A1:C5");
  if (!controller.readOnly) controller.setCellValue({ row: 0, col: 0 }, "42");
  if (controller.canExport) controller.exportXlsx();
}
```

`setCellValue` requires an editable workbook. `exportXlsx()` downloads the current workbook in the browser.

Change options or load another workbook with `setOptions`:

```js
viewer.setOptions({ isDark: true, readOnly: true });
viewer.setOptions({ src: "/another-workbook.xlsx", fileName: undefined });
```

Supplying `file` clears the previous `src`, and supplying `src` clears the previous `file`. To release the viewer when it is no longer needed:

```js
unsubscribe();
viewer.destroy();
```

`createXlsxViewer` returns an object with these methods:

- `getController()` returns the latest `XlsxViewerController`, or `null` when no controller is available or after destruction. Read it again when invoking commands so you have the current state and methods.
- `subscribe(listener)` reports controller updates, including loading, errors, sheets, selection, zoom, edits, and the existing controller commands. If a controller is already available, the listener is called immediately. Updates can be coalesced. The return value unsubscribes the listener.
- `setOptions(options)` merges new options into the current ones. Supplying `file` clears `src`, and supplying `src` clears `file`. `file` accepts an `ArrayBuffer`; convert a browser `File` with `await file.arrayBuffer()`.
- `reload()` discards the current viewer state and reloads the current source.
- `destroy()` unmounts the viewer, releases its resources, and removes listeners. It is safe to call more than once.

`subscribe` also exposes `isLoadDeferred` and `continueDeferredLoad()` for deferred workbooks.

`XlsxVanillaViewerOptions` supports the non-React viewer options, including `getCellStyle`, form-control events, themes, worker settings, and highlighting. `toolbar`, loading/empty/error states, and the `renderActiveCellOverlay`, `renderChartLoading`, `renderContextMenu`, `renderFormControl`, and `renderImage` callbacks accept text or DOM nodes. Their vanilla callback props omit React's `defaultNode`. React-specific custom scroller, image-selection, and table-header render callbacks are not available through this entry point.

The vanilla entry also exports `setWasmSource`, `initWasm`, `parseXlsxRange`, and workbook types. The same WASM, worker, font, and chart assets need to be served as in the React integration.

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

## Merging Cells

The controller and `useXlsxViewerEditing()` expose `mergeSelection(mode?)` and `unmergeSelection()`. The merge modes are:

| Mode | Behavior |
| --- | --- |
| `"cells"` (default) | Merge the entire selection and preserve the anchor's alignment. |
| `"center"` | Merge the entire selection and persist horizontal center alignment. |
| `"across"` | Merge each selected row separately, preserving each row's anchor alignment. |

The playground Merge menu and the editable viewer's context menu expose all three modes. Merge and alignment changes participate in undo/redo and persist in XLSX exports. Unmerge works when selecting any cell within a merged range.

General alignment follows the underlying value type: numbers align right, text aligns left, and booleans and errors align center. Explicit horizontal and vertical alignment is preserved for both merged and ordinary cells.

## Form Control Editing

When `readOnly={false}`, Duke-backed checkboxes, option buttons, dropdowns, list boxes, scrollbars, and spinners are interactive. Changes update linked cells immediately, participate in undo/redo, and are written by `exportXlsx()` / `download()`.

Buttons do not execute Excel macros. Provide `onFormControlAction` to handle a button activation in your application. Labels and group boxes are display-only.

```tsx
<XlsxViewer
  file={file}
  onFormControlAction={({ control }) => {
    console.log("Activated", control.name);
  }}
  onFormControlChange={({ control, previousControl }) => {
    console.log(previousControl.state, "->", control.state);
  }}
/>
```

The controller also exposes `addFormControl(input, sheetIndex?)`, `updateFormControl(controlIndex, patch, sheetIndex?)`, `removeFormControl(controlIndex, sheetIndex?)`, and `getFormControlItems(controlIndex, sheetIndex?)`. `XlsxFormControlInput` uses the same `XlsxImageAnchor` union as images and charts, with an optional `editAs` value for two-cell anchors. Captions accept either a string or rich text runs:

```ts
const controlIndex = controller.addFormControl({
  anchor: {
    kind: "two-cell",
    from: { col: 1, row: 1, colOffsetEmu: 0, rowOffsetEmu: 0 },
    to: { col: 3, row: 2, colOffsetEmu: 0, rowOffsetEmu: 0 }
  },
  editAs: "twoCell",
  kind: {
    kind: "checkbox",
    caption: {
      runs: [{ text: "Include forecast", font: { bold: true, name: "Aptos", size: 10 } }]
    },
    state: "checked"
  }
});
```

### Migrating To 0.15

- `XlsxFormControlAnchor` was removed. Pass an `XlsxImageAnchor` to `XlsxFormControlInput.anchor` instead.
- Read controls now expose `caption` as `XlsxFormControlCaption`; use `label` when you only need flattened display text.
- Mutation captions accept either a plain string or `XlsxFormControlCaption` rich text runs.
- `firstInGroup` is read-only. Duke recomputes option-button grouping when controls are written.

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

## Range Highlighting

Pass an A1 reference or an array of references to `highlightedRanges` to draw blue bounding boxes with a thin border matching the selection outline and a subtle fill. The highlights are separate from the current selection and do not change workbook formatting. It works in both canvas and DOM rendering, including frozen panes, hidden rows/columns, merged cells, and zoom.

```tsx
<XlsxViewer
  file={file}
  highlightedRanges={["'Sheet 1'!$B$3:$E$8", "'Sheet 1'!G3:J8", "'Sheet 2'!B10:D14"]}
  autoScrollToHighlightedRanges={true}
  highlightClassName="my-range-highlight"
/>
```

`autoScrollToHighlightedRanges` defaults to `true`: a new set of ranges reveals the first range and switches to its worksheet. Other ranges remain highlighted on their respective sheets. Set it to `false` to leave both the current scroll position and active sheet unchanged. Navigation happens once per set of ranges or auto-scroll option change, so manual scrolling remains free afterward. A range larger than the viewport reveals its top-left portion; whole-row and whole-column references scroll only the relevant axis.

Set `highlightedRanges={null}` or `highlightedRanges={[]}` to clear all highlights. Omit the prop to manage highlighting through the controller instead. If any prop reference is invalid or unresolved, the previous highlights are cleared. Equivalent arrays do not repeat navigation on parent re-renders. An unqualified reference is bound to the active worksheet when applied, and stays on that worksheet when the user switches tabs.

### Controller API

The same API is available through `useXlsxViewer()` and `useXlsxViewerController()`:

```tsx
const accepted = controller.highlightRanges(["B3:E8", "G3:J8"], { autoScroll: false });
controller.highlightRanges("'Sheet 1'!A40:D48");
controller.highlightRanges("B3:E8", { sheetIndex: 1 });
controller.clearHighlightedRanges();
```

`highlightRanges(references, options?)` accepts a string or a readonly array of strings and returns `true` when every reference resolves to a loaded, visible worksheet. It returns `false` for invalid references, unavailable worksheets, or a workbook that is still loading, leaving all current highlights and navigation unchanged. Validation is atomic; a partially valid array is never applied. `autoScroll` defaults to `true`; `sheetIndex` is a zero-based index into `controller.sheets` and applies only to unqualified references. Each successful call replaces all previous highlights; an empty array clears them. Calling again can reveal the same first range again. `controller.highlightedRanges` is an array of parsed references with their resolved sheet names and workbook sheet indexes. Only its first entry can have `autoScroll: true`. Loading another workbook clears the array.

### CSS Customization

Style `.react-xlsx-range-highlight` globally, or use `highlightClassName` for a scoped override. Appearance is CSS in both renderers, so ordinary CSS properties work without `!important`:

```css
.my-range-highlight {
  border: 1px solid #0284c7;
  background: rgb(2 132 199 / 10%);
  border-radius: 3px;
}
```

The defaults also accept inherited `--xlsx-highlight-border`, `--xlsx-highlight-fill`, and `--xlsx-highlight-radius` custom properties on the viewer or any ancestor. Position and dimensions are managed by the viewer; the overlay ignores pointer events. Use `[data-xlsx-range-highlight="1"]` to style a particular range by its zero-based index in the input array. A range crossing frozen panes may have several clipped overlay elements, each carrying the same class and index.

### Parsing References

`parseXlsxRange(reference)` validates a reference without loading a workbook. It returns `{ kind, range: { start, end }, sheetName }` with normalized, zero-based coordinates, or `null` for invalid input. `kind` is `"cells"`, `"rows"`, or `"columns"`; `sheetName` is `null` for an unqualified reference.

Supported references include single cells (`A1`), rectangles (`B2:E8`), absolute or mixed references (`$B2:E$8`), entire rows (`2:8`), entire columns (`B:E`), and optional worksheet prefixes. Sheet names with spaces or punctuation must be single-quoted; embedded apostrophes are doubled (`'Owner''s Data'!A1`). Letter case, reversed endpoints, surrounding whitespace, and an optional leading `=` are accepted. Coordinates are checked against [Excel's worksheet limits](https://support.microsoft.com/en-us/excel/excel-specifications-and-limits).

Each string must contain one rectangular A1 reference, not named ranges, structured table references, unions, R1C1 notation, external workbooks, or references spanning multiple worksheets. Whole-row/column highlights cover the displayed grid and extend as it grows.

## Examples

See `@extend-ai/react-xlsx` in action in these live [Extend UI](https://www.extend.ai/ui) examples. Click a screenshot to open the demo and integration docs.

### [XLSX Viewer](https://www.extend.ai/ui/docs/components/xlsx-viewer)

Explore workbook navigation, sheet tabs, zoom, and charts.

[![XLSX viewer with worksheet data, a column chart, and sheet tabs](https://raw.githubusercontent.com/extend-hq/react-xlsx/main/docs/readme-assets/extend-ui-xlsx-viewer.png)](https://www.extend.ai/ui/docs/components/xlsx-viewer)

### [XLSX Editor](https://www.extend.ai/ui/docs/components/xlsx-editor)

Try the experimental editor with cell formatting, formulas, undo/redo, and export.

[![XLSX editor with formatting controls and a styled worksheet](https://raw.githubusercontent.com/extend-hq/react-xlsx/main/docs/readme-assets/extend-ui-xlsx-editor.png)](https://www.extend.ai/ui/docs/components/xlsx-editor)

## Context Menus

Right-click a worksheet cell, row header, column header, or the top-left corner to open the context menu. Right-clicking inside a selection preserves the range. Shift+F10 or the context-menu key opens it at the active cell; arrow keys navigate and Escape dismisses it.

Viewer mode (`readOnly`) offers Copy and selection actions. Editor mode also offers Cut, Paste, Edit cell, Clear contents, Undo/Redo, Merge/Unmerge, and sorting when the target belongs to a table. Actions use the existing controller and undo history. Row and column selections span the displayed grid, matching header selection. Worker-backed viewers copy tab-separated text; editable workbooks also preserve structured cell data when the browser supports custom clipboard formats. Clipboard commands require browser clipboard access; failed operations display an error and Cut only clears cells after copying succeeds.

Customize the existing menu with `contextMenuClassName`, `contextMenuStyle`, and CSS variables, or replace its contents with `renderContextMenu`:

```tsx
<XlsxViewer
  file={file}
  contextMenuClassName="spreadsheet-menu"
  contextMenuStyle={{ borderRadius: 6 }}
  renderContextMenu={({ context, actions, runAction, closeMenu, busy }) => (
    <>
      {actions.map((action) => (
        <button
          key={action.id}
          type="button"
          role="menuitem"
          disabled={busy || action.disabled}
          onClick={() => void runAction(action.id)}
        >
          {action.label}
        </button>
      ))}
      <button type="button" role="menuitem" onClick={() => {
        inspectSelection(context.selection);
        closeMenu();
      }}>
        Inspect selection
      </button>
    </>
  )}
/>
```

The hook receives `context` (target kind, cell, selection, sheet index/name, and read-only state), `actions`, `runAction`, `closeMenu`, viewport `position`, `isDark`, `busy`, `error`, and `defaultNode`. Return `defaultNode` to keep built-in buttons while adding your own content. The viewer owns the portal, viewport-edge positioning, outside-click dismissal, and keyboard navigation for elements with `role="menuitem"`; return menu contents without another portal. `runAction` returns `Promise<boolean>` and rejects unavailable or disabled commands. Handle custom commands in your own callback and call `closeMenu` afterward.

Default styling follows the document theme. Override `.react-xlsx-context-menu`, `.react-xlsx-context-menu-item`, `.react-xlsx-context-menu-separator`, and `.react-xlsx-context-menu-shortcut`, or set `--xlsx-context-menu-background`, `--xlsx-context-menu-color`, `--xlsx-context-menu-border`, `--xlsx-context-menu-hover`, `--xlsx-context-menu-muted`, and `--xlsx-context-menu-danger` through `contextMenuStyle`. The menu is portaled to the document body, so use its class rather than a viewer-descendant selector. Set `enableContextMenu={false}` to retain the browser's native context menu.

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

Use `renderActiveCellOverlay` to render your own content over the active cell. The returned node sits in a container that exactly covers the cell (the full merge for merged cells) and moves with scrolling and zoom, so you can attach classes, data attributes, or a ref and measure it with `getBoundingClientRect()`. The container ignores pointer events; set `pointerEvents: "auto"` on your own elements to make them interactive.

```tsx
<XlsxViewer
  file={file}
  renderActiveCellOverlay={({ cell }) => (
    <div
      ref={activeCellRef}
      className="active-cell-anchor"
      data-active-cell={`${cell.row}:${cell.col}`}
      style={{ height: "100%", width: "100%" }}
    />
  )}
/>
```

Use `renderFormControl` to replace the built-in worksheet widgets. The callback receives the parsed `control`, resolved `items` and `label`, calculated `style`, read-only state, and `activate`, `setState`, `setSelected`, and `setValue` helpers. Apply `style` to the custom root and call `stopPropagation` from pointer/click handlers. The helpers keep Duke mutations, linked cells, undo/redo, export, and `onFormControlChange` in sync.

```tsx
<XlsxViewer
  file={file}
  renderFormControl={({ checked, control, disabled, label, setState, stopPropagation, style }) => {
    if (control.kind === "checkbox" || control.kind === "radio") {
      return (
        <label onPointerDown={stopPropagation} style={style}>
          <input
            checked={checked}
            disabled={disabled}
            onChange={(event) => setState(event.currentTarget.checked ? "checked" : "unchecked")}
            type={control.kind === "radio" ? "radio" : "checkbox"}
          />
          {label}
        </label>
      );
    }

    return <div style={style}>{label}</div>;
  }}
/>
```

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
- `headerBackgroundColor?: string`
- `headerTextColor?: string`
- `rounded?: boolean`
- `showDefaultToolbar?: boolean`
- `toolbar?: React.ReactNode | ((controller: XlsxViewerController) => React.ReactNode)`
- `experimentalCanvas?: boolean`
- `enableGestureZoom?: boolean`
- `enableCanvasSelectionAnimation?: boolean`
- `allowResizeInReadOnly?: boolean`
- `highlightedRanges?: string | readonly string[] | null`
- `autoScrollToHighlightedRanges?: boolean`
- `highlightClassName?: string`
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
- `renderFormControl?: (props: XlsxFormControlRenderProps) => React.ReactNode`
- `renderTableHeaderMenu?: (props: XlsxTableHeaderMenuRenderProps) => React.ReactNode`
- `renderActiveCellOverlay?: (props: XlsxActiveCellOverlayRenderProps) => React.ReactNode`
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
- `XlsxFormControl`, `XlsxFormControlInput`, `XlsxFormControlCaption`, `XlsxFormControlCaptionInput`, `XlsxFormControlCaptionRun`, `XlsxFormControlRenderProps`
- `XlsxTable`, `XlsxTableColumn`, `XlsxTableHeaderMenuRenderProps`
- `XlsxActiveCellOverlayRenderProps`
- `XlsxWorkbookTab`, `XlsxCellAddress`, `XlsxCellRange`
- `XlsxParsedRange`, `XlsxRangeHighlight`, `XlsxHighlightRangeOptions`

## License

See the repository license for usage terms.
