import { readFile } from "node:fs/promises";
import type { Options } from "tsup";

export function atlasJsonPlugin(): NonNullable<Options["esbuildPlugins"]>[number] {
  return {
    name: "atlas-json",
    setup(build) {
      build.onLoad({ filter: /[\\/]node_modules[\\/](?:us-atlas|world-atlas)[\\/].*\.json$/ }, async ({ path }) => {
        const json = JSON.stringify(JSON.parse(await readFile(path, "utf8")));
        return {
          contents: `export default /* @__PURE__ */ JSON.parse(${JSON.stringify(json)});`,
          loader: "js"
        };
      });
    }
  };
}
