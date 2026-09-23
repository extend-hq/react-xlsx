import { defineConfig } from "tsup";
import { atlasJsonPlugin } from "./build/atlas-json.ts";

const external = ["react", "react-dom"];
const noExternal = ["us-atlas", "world-atlas"];

export default defineConfig([
  {
    clean: true,
    dts: true,
    entry: ["src/index.ts"],
    esbuildPlugins: [atlasJsonPlugin()],
    external,
    format: ["esm", "cjs"],
    noExternal,
    skipNodeModulesBundle: true,
    sourcemap: true
  },
  {
    clean: false,
    entry: ["src/xlsx-worker.ts"],
    esbuildPlugins: [atlasJsonPlugin()],
    external,
    format: ["esm"],
    noExternal,
    outDir: "dist",
    skipNodeModulesBundle: true,
    sourcemap: true
  }
]);
