# Office Charts map renderer

Generated from the shared Office Charts Rust renderer. Maps load this module lazily and do not use JavaScript projection or geometry libraries. Saved Office geometry takes precedence; absent geometry can use the optional offline country and US-state atlas. Other uncached subdivisions remain unavailable and are described in the SVG.

This checked-in module allows the viewer to use the local renderer without depending on an unpublished package release. The compiled source and asset digests are recorded in provenance.json, and both required license notices are retained. Do not edit generated JavaScript or WASM manually.

Refresh with `node scripts/sync-office-maps.mjs /path/to/extend-office` after running `npm run build:browser` there. The build copies the raw WASM fallback and license notices into the package; modern browsers normally load the gzip payload without a network request.
