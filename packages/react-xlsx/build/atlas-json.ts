import { readFile } from "node:fs/promises";
import type { Options } from "tsup";

type ArcReference = number | ArcReference[];

type PolygonAtlas = {
  arcs: number[][][];
  objects: Record<string, { geometries: Array<{ arcs: ArcReference }> }>;
};

function selectPolygonLayer(atlas: PolygonAtlas, name: string) {
  const layer = atlas.objects[name];
  const arcs: number[][][] = [];
  const indices = new Map<number, number>();

  function remap(reference: ArcReference): ArcReference {
    if (Array.isArray(reference)) return reference.map(remap);
    const source = reference < 0 ? ~reference : reference;
    let index = indices.get(source);
    if (index === undefined) {
      index = arcs.length;
      indices.set(source, index);
      arcs.push(atlas.arcs[source]);
    }
    return reference < 0 ? ~index : index;
  }

  const geometries = layer.geometries.map(geometry => ({ ...geometry, arcs: remap(geometry.arcs) }));
  return { ...atlas, objects: { [name]: { ...layer, geometries } }, arcs };
}

export function atlasJsonPlugin(): NonNullable<Options["esbuildPlugins"]>[number] {
  return {
    name: "atlas-json",
    setup(build) {
      build.onLoad({ filter: /[\\/]node_modules[\\/](?:us-atlas|world-atlas)[\\/].*\.json$/ }, async ({ path, with: attributes }) => {
        const atlas = JSON.parse(await readFile(path, "utf8"));
        const json = JSON.stringify(attributes.layer ? selectPolygonLayer(atlas, attributes.layer) : atlas);
        return {
          contents: `export default /* @__PURE__ */ JSON.parse(${JSON.stringify(json)});`,
          loader: "js"
        };
      });
    }
  };
}
