import { readFileSync } from "node:fs";
import { officeArtSourceDigest, sha256 } from "./office-art-provenance.mjs";

const root = new URL("../packages/react-xlsx/src/office-art/generated/", import.meta.url);
const provenance = JSON.parse(readFileSync(new URL("provenance.json", root), "utf8"));
if (provenance.sourceDigest !== officeArtSourceDigest()) throw new Error("Office artwork bindings need rebuilding with pnpm build:office-art.");
for (const [name, digest] of Object.entries(provenance.outputs)) {
  if (sha256(readFileSync(new URL(name, root))) !== digest) throw new Error("Office artwork build output failed its integrity check.");
}
