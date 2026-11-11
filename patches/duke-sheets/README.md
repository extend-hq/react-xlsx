# Bundled spreadsheet parser

The package builds DukeSheets from upstream commit `fec1cd04dbefa5eb65a41d29399ca4a7edc220f2` with the checked-in patch. No fork or upstream release is required. Run `pnpm build:duke-sheets` with Rust's `wasm32-unknown-unknown` target and `wasm-pack` installed. The script creates a temporary checkout, verifies and applies the patch, builds with the locked dependencies, and records output hashes. Ordinary package builds verify those hashes and bundle the generated runtime.

The patch exposes rich-text runs, preserves the archive entries inflated by the native reader, and provides a Rust XML decoder for viewer metadata. ZIP inflation and XML tokenization run in the worker. Existing feature projections for styles, conditional formatting, charts, drawings, images, controls, and round-trip preservation consume the decoded tree in that worker. They remain TypeScript so that this change retains the viewer's existing interpretation of those features.

Editable workbooks cross the worker boundary as a binary model snapshot. Restoring that model does not reopen or parse the source spreadsheet. Snapshots include the native model and round-trip style/theme state; they are internal to this exact runtime build and are not a persistent interchange format. Export runs in the worker, including merging original drawing parts and edited XML. Browser rendering and editing of the hydrated model remain on the main thread.

To update upstream, update the commit in `scripts/duke-sheets-provenance.mjs`, rebase the patch, rebuild, and validate workbook features, rich-text rendering, editing, undo, and export before publishing. Generated bindings, WASM, provenance, and the upstream license are checked in together.
