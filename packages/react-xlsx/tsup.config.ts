import { defineConfig, type Options } from "tsup";

const external = ["react", "react-dom"];
const officeCharts = ["@extend-ai/rust-office-charts", "@extend-ai/office-fonts"];
const virtualizer = ["@tanstack/react-virtual", "@tanstack/virtual-core"];

function sharedCorePlugin(): NonNullable<Options["esbuildPlugins"]>[number] {
  return {
    name: "shared-xlsx-core",
    setup(build) {
      build.onResolve({ filter: /^@extend-ai\/react-xlsx\/(?:internal|vanilla)$/ }, (args) => ({ path: args.path, external: true }));
      build.onResolve({ filter: /^\.\/(wasm|errors)$/ }, (args) => {
        if (/\/(controller\.tsx|worker-client\.ts)$/.test(args.importer)) {
          return { path: "@extend-ai/react-xlsx/internal", external: true };
        }
        return undefined;
      });
    }
  };
}

export default defineConfig([
  {
    clean: true,
    dts: true,
    entry: { vanilla: "src/vanilla-entry.ts", internal: "src/internal.ts" },
    format: ["esm", "cjs"],
    noExternal: [...officeCharts, ...virtualizer, "preact", "react", "react-dom"],
    loader: { ".woff2": "dataurl" },
    esbuildOptions(options) {
      options.alias = { ...options.alias, react: "preact/compat", "react-dom": "preact/compat" };
    },
    shims: true,
    skipNodeModulesBundle: true,
    splitting: true,
    sourcemap: true
  },
  {
    clean: false,
    dts: true,
    entry: ["src/index.ts"],
    esbuildPlugins: [sharedCorePlugin()],
    external: [...external, "@extend-ai/react-xlsx/vanilla", "@extend-ai/react-xlsx/internal"],
    format: ["esm", "cjs"],
    noExternal: [...officeCharts, ...virtualizer],
    loader: { ".woff2": "dataurl" },
    shims: true,
    skipNodeModulesBundle: true,
    splitting: true,
    sourcemap: true
  },
  {
    clean: false,
    entry: ["src/xlsx-worker.ts"],
    external,
    format: ["esm"],
    noExternal: ["@dukelib/sheets-wasm", "fflate"],
    outDir: "dist",
    platform: "browser",
    skipNodeModulesBundle: false,
    splitting: false,
    sourcemap: true
  }
]);
