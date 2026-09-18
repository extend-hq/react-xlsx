import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { fileURLToPath, pathToFileURL } from "node:url";
import { build, type Options } from "tsup";
import ts from "typescript";
import configuration from "../tsup.config.ts";

test("atlas bundles preserve geography with a bounded JavaScript syntax tree in ESM and CJS", async () => {
  const temporary = await mkdtemp(join(tmpdir(), "xlsx-atlas-"));
  try {
    const modules = ["us-atlas/counties-albers-10m.json", "world-atlas/countries-50m.json"];
    const paths = modules.map((name) => fileURLToPath(import.meta.resolve(name)));
    const expected = await Promise.all(paths.map(async (path) => JSON.parse(await readFile(path, "utf8"))));
    const entry = join(temporary, "input.js");
    await writeFile(entry, [
      ...paths.map((path, index) => `import atlas${index} from ${JSON.stringify(path)};`),
      "export default [atlas0, atlas1];"
    ].join("\n"));

    assert.ok(Array.isArray(configuration));
    const buildOptions: Options = {
      ...configuration[0],
      config: false,
      entry: { atlas: entry },
      outDir: temporary,
      outExtension: ({ format }) => ({ js: format === "esm" ? ".mjs" : ".cjs" }),
      clean: false,
      dts: false,
      silent: true
    };
    await build(buildOptions);

    const loadCommonJs = createRequire(import.meta.url);
    for (const extension of ["mjs", "cjs"]) {
      const output = join(temporary, `atlas.${extension}`);
      const source = await readFile(output, "utf8");
      const tree = ts.createSourceFile(output, source, ts.ScriptTarget.Latest, false, ts.ScriptKind.JS);
      let nodeCount = 0;
      function countNodes(node: ts.Node) {
        nodeCount += 1;
        ts.forEachChild(node, countNodes);
      }
      countNodes(tree);
      assert.ok(nodeCount < 1000, `${extension} bundle expands geographic data into ${nodeCount} syntax nodes`);
      const module = extension === "mjs" ? await import(pathToFileURL(output).href) : loadCommonJs(output);
      assert.deepEqual(module.default, expected);
    }

    await writeFile(entry, [
      ...paths.map((path, index) => `import atlas${index} from ${JSON.stringify(path)};`),
      "export default null;"
    ].join("\n"));
    await build(buildOptions);
    for (const extension of ["mjs", "cjs"]) {
      const source = await readFile(join(temporary, `atlas.${extension}`), "utf8");
      assert.ok(source.length < 10000, `${extension} bundle retains unused atlas data`);
    }
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
});
