# Changelog

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
