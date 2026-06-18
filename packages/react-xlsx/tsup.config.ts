import { defineConfig } from "tsup";

const external = ["react", "react-dom"];
const noExternal = ["us-atlas", "world-atlas"];

export default defineConfig([
  {
    clean: true,
    dts: true,
    entry: ["src/index.ts"],
    external,
    format: ["esm", "cjs"],
    noExternal,
    skipNodeModulesBundle: true,
    sourcemap: true
  },
  {
    clean: false,
    entry: ["src/xlsx-worker.ts"],
    external,
    format: ["esm"],
    noExternal,
    outDir: "dist",
    skipNodeModulesBundle: true,
    sourcemap: true
  }
]);
