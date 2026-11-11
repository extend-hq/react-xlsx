# Office picture bindings

This adapter delegates XLSX picture appearance to the published `rust-office-art = 0.1.8` crate. It supplies the workbook's color resolver and exposes the shared SVG renderer and raster normalization to the browser. It does not implement a second picture renderer.

Run `pnpm build:office-art` after changing the adapter or its locked dependencies. Rust, the `wasm32-unknown-unknown` target, and `wasm-pack` are required only for rebuilding the bindings. The normal package build verifies the checked-in JavaScript, embedded compressed WASM, source digest, and third-party notices without requiring a Rust toolchain. `pnpm build:office-art --offline` uses cached dependencies.

The browser loads the compressed module only for workbooks containing DrawingML pictures. All parsed pictures go through the shared renderer, including its unchanged-image result. Unsupported decorations retain the source image, expose `pictureRenderStatus: "unsupported"`, and emit a generic warning. The original media and drawing XML remain in the workbook archive for export. SVG paint bounds are separate from image anchors so borders, rotations, shadows, and reflections do not change editing geometry.

The adapter canonicalizes spreadsheet picture namespaces and accounts for Excel's quarter-turn anchor dimensions. Canvas, DOM, and thumbnail drawing consume the same rendered source and paint bounds.

General shapes, connectors, SmartArt, and legacy binary spreadsheet artwork remain outside this picture migration. The shared crate's unsupported effects (including 3D scenes and arbitrary effect graphs) remain unsupported; a fallback is not a claim of visual parity.

The explicit `skipXmlParsing` option retains its existing opt-out behavior and bypasses DrawingML picture decoration.

Native Excel comparison identified upstream differences in preset border dashes and tiled-image DPI handling. Keep those fixes in `rust-office-art`: this adapter uses the published crate without patches or effect workarounds. Update the locked dependency and rebuild the bindings when upstream fixes are released.
