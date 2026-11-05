import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";

export const sha256 = (bytes) => createHash("sha256").update(bytes).digest("hex");
export function officeArtSourceDigest() {
  return sha256(Buffer.concat([
    "rust/office-art-wasm/Cargo.toml",
    "rust/office-art-wasm/Cargo.lock",
    "rust/office-art-wasm/src/lib.rs",
    "scripts/build-office-art.mjs",
    "scripts/office-art-provenance.mjs"
  ].flatMap((path) => [Buffer.from(`${path}\0`), readFileSync(new URL(`../${path}`, import.meta.url))])));
}
