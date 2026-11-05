import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { gzipSync } from "node:zlib";
import { fileURLToPath } from "node:url";
import { homedir } from "node:os";
import { dirname, join } from "node:path";
import { officeArtSourceDigest, sha256 } from "./office-art-provenance.mjs";

const root = fileURLToPath(new URL("../", import.meta.url));
const output = fileURLToPath(new URL("../packages/react-xlsx/src/office-art/generated/", import.meta.url));
mkdirSync(output, { recursive: true });
const offline = process.argv.includes("--offline") ? ["--offline"] : [];
const sourceDigest = officeArtSourceDigest();
const result = spawnSync("wasm-pack", ["build", "rust/office-art-wasm", "--target", "web", "--release", "--no-pack", "--out-dir", output, "--out-name", "office_art", "--", "--locked", ...offline], {
  cwd: root,
  stdio: "inherit",
  env: { ...process.env, CARGO_ENCODED_RUSTFLAGS: `--remap-path-prefix=${homedir()}=/home/builder\x1f--remap-path-prefix=${root}=.` }
});
if (result.error) throw result.error;
if (result.status !== 0) process.exit(result.status ?? 1);
const bytes = readFileSync(new URL("../packages/react-xlsx/src/office-art/generated/office_art_bg.wasm", import.meta.url));
writeFileSync(`${output}/payload.js`, `export default "${gzipSync(bytes, { level: 9 }).toString("base64")}";\n`);
writeFileSync(`${output}/payload.d.ts`, "declare const payload: string;\nexport default payload;\n");
for (const name of [".gitignore", "package.json", "office_art_bg.wasm", "office_art_bg.wasm.d.ts"]) rmSync(`${output}/${name}`, { force: true });
const metadata = spawnSync("cargo", ["metadata", "--manifest-path", "rust/office-art-wasm/Cargo.toml", "--locked", "--format-version", "1", ...offline], { cwd: root, encoding: "utf8", maxBuffer: 16 * 1024 * 1024 });
if (metadata.status !== 0) throw new Error(metadata.stderr || "Unable to read Rust package metadata.");
const packages = JSON.parse(metadata.stdout).packages;
const notices = packages.filter((pkg) => pkg.source).map((pkg) => {
  const directory = dirname(pkg.manifest_path);
  const paths = [...new Set([pkg.license_file, "LICENSE", "LICENSE-MIT", "LICENSE-APACHE", "LICENSE.txt", "LICENSE.md", "COPYING"].filter(Boolean))];
  const licenses = paths.map((name) => join(directory, name)).filter(existsSync).map((path) => readFileSync(path, "utf8").replace(/[ \t]+$/gm, ""));
  return `${pkg.name} ${pkg.version}\n${pkg.license ?? ""}\n${pkg.repository ?? ""}\n\n${licenses.join("\n\n")}`;
});
writeFileSync(`${output}/THIRD-PARTY-NOTICES.txt`, notices.join("\n\n--------------------\n\n").trimEnd() + "\n");
if (officeArtSourceDigest() !== sourceDigest) throw new Error("Office artwork sources changed during compilation.");
const names = ["office_art.js", "office_art.d.ts", "payload.js", "payload.d.ts", "THIRD-PARTY-NOTICES.txt"];
writeFileSync(`${output}/provenance.json`, JSON.stringify({ crate: "rust-office-art", version: "0.1.1", sourceDigest, wasmSha256: sha256(bytes), outputs: Object.fromEntries(names.map((name) => [name, sha256(readFileSync(`${output}/${name}`))])) }, null, 2) + "\n");
