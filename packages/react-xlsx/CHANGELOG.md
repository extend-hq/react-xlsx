# Changelog

## 0.22.1

- Match Excel's opaque data-bar gradients, cell inset, and default length limits in the grid and thumbnails while preserving explicit length settings.
- Format imported ISO date cells using their number formats and workbook date system.
- Keep table header menus within the visible grid and share their styling with context menus.

## 0.22.0

- Add worksheet-backed pivot tables with editing, refresh, filtering, grouping, calculations, drill-through, PivotCharts, undo/redo, and XLSX import/export.
- Expose React pivot editor hooks and customizable field panes and header menus for React and vanilla JavaScript.
- Use workers by default for parsing, charts, export, and read-only queries, with automatic main-thread fallback and a complete `useWorker: false` opt-out.
- Keep worksheet images available after worker failures and preserve edited workbook data during export fallback.
- Update the bundled spreadsheet and Office rendering runtimes.
- Breaking: `getXlsxBytes()` and `exportXlsx()` now return promises and must be awaited. Self-hosted parser WASM must come from the same package version; shared table header menus also receive pivot headers.

## 0.21.0

- Add opt-in custom cell editor overlays and selection popovers for React and vanilla JavaScript, with caller-owned UI, styling, animations, and callbacks.
- Expose editor draft replacement, preview, discard, commit, focus, and cancellation commands with abort signals for asynchronous integrations.
- Add guarded custom context-menu commits and batch cell styling with a single undo step.
- Redesign the playground with Coss UI, formula and heatmap previews, and an always-visible read-only switch.

## 0.20.12

- Use data-type-neutral ascending and descending sort labels.
- Localize built-in menus, toolbar controls, status text, and accessibility labels with eight bundled languages and English fallback.
- Add partial `messages` overrides to the React and vanilla viewers and shared controllers.

## 0.20.11

- Restore table sort indicators along with row order during undo and redo.
- Preserve cell selections when opening table sort menus or sorting, including custom menus and frozen headers in the canvas renderer.

## 0.20.10

- Fix Next.js Turbopack production builds while keeping OfficeArt bundled and enabled by default.

## 0.20.8

- Add formula-aware cell editing and highlight referenced ranges while editing.
- Render drawings across frozen pane boundaries and refine chart selection and source range highlights.
- Improve canvas table header rendering and menu interaction.

## 0.20.7

- Use the published OfficeArt 0.1.3 picture renderer and chart browser 0.3.1 release.
- Render chart picture fills, picture markers, drawing overlays and supported artwork effects through the shared Rust renderer.
- Preserve chart selection when decorative effects are present, and improve radar and stock markers.
- Correct picture border dash spacing and image tile density through the shared renderer.

## 0.20.5

- Use the bundled spreadsheet font fallback for Office chart titles and labels when native fonts are unavailable.
- Update Rust Office Charts to improve title font defaults, wrapping, manual placement, and radial gradient fidelity.

## 0.20.4

- Use the updated Rust Office renderer for additional chart families and filled maps, preserving series and point selection.
- Remove JavaScript geographic rendering dependencies and bundle offline map boundaries.
- Improve chart data hydration, chartsheet rendering, and surface rendering; bundle self-contained chart types.

## 0.20.3

- Bundle the Rust Office chart runtime and WASM so consumers install only registry dependencies.

## 0.20.2

- Use the Rust Office crates by Extend and refresh the bundled WebAssembly.
