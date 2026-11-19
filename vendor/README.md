The viewer takes the chart renderer from the versioned Duke Office browser release on GitHub. The package manifest pins the release asset and the pnpm lockfile records its integrity. The adjacent provenance record identifies the source revision, archive digest and both WASM binaries bundled into the viewer.

To try a local renderer build, run `node scripts/vendor-office-charts.mjs <renderer-checkout>`, refresh the lockfile and rebuild the viewer. For a release, pin the published asset and verify its core and full WASM digests against the tested build.
