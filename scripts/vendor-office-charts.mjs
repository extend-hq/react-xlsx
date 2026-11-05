import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import { copyFileSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const repository = resolve(dirname(fileURLToPath(import.meta.url)), "..");
if (!process.argv[2]) throw new Error("Pass the path to the renderer source checkout after building its browser package.");
const source = resolve(process.argv[2]);
const sourcePackage = JSON.parse(readFileSync(join(source, "package.json"), "utf8"));
if (sourcePackage.name !== "@extend-ai/rust-office-charts") throw new Error("Unexpected renderer package.");
const run = (command, args, cwd = source, inherit = false) => {
  const result = spawnSync(command, args, { cwd, encoding: "utf8", stdio: inherit ? "inherit" : "pipe" });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error(`${command} failed: ${result.stderr ?? result.status}`);
  return result.stdout?.trim() ?? "";
};
const hash = (bytes, encoding = "hex", algorithm = "sha256") => createHash(algorithm).update(bytes).digest(encoding);

run(process.execPath, ["browser/check.mjs"], source, true);
const temporary = mkdtempSync(join(tmpdir(), "office-charts-vendor-"));
try {
  const [packed] = JSON.parse(run("npm", ["pack", "--ignore-scripts", "--json", "--pack-destination", temporary, "--cache", join(temporary, "npm-cache")]));
  const archive = join(temporary, packed.filename);
  const bytes = readFileSync(archive);
  const digest = hash(bytes);
  const filename = `rust-office-charts-${sourcePackage.version}-${digest.slice(0, 16)}.tgz`;
  const vendor = join(repository, "vendor");
  mkdirSync(vendor, { recursive: true });
  copyFileSync(archive, join(vendor, filename));
  const provenance = {
    name: sourcePackage.name,
    version: sourcePackage.version,
    archive: filename,
    sha256: digest,
    integrity: `sha512-${hash(bytes, "base64", "sha512")}`,
    sourceRevision: run("git", ["rev-parse", "HEAD"]),
    sourceHasUncommittedChanges: Boolean(run("git", ["status", "--porcelain"])),
    sourceDigest: readFileSync(join(source, "browser/generated/sources.sha256"), "utf8").trim(),
    coreWasmSha256: hash(readFileSync(join(source, "browser/generated/office_charts_bg.wasm"))),
    fullWasmSha256: hash(readFileSync(join(source, "browser/generated/full/office_charts_bg.wasm")))
  };
  writeFileSync(join(vendor, filename.replace(/\.tgz$/, ".json")), JSON.stringify(provenance, null, 2) + "\n");
  const manifest = join(repository, "packages/react-xlsx/package.json");
  const contents = JSON.parse(readFileSync(manifest, "utf8"));
  contents.devDependencies[sourcePackage.name] = `file:../../vendor/${filename}`;
  writeFileSync(manifest, JSON.stringify(contents, null, 2) + "\n");
  process.stdout.write(`Vendored ${filename}\nRun pnpm install --lockfile-only --offline --ignore-scripts, then rebuild the viewer package.\n`);
} finally {
  rmSync(temporary, { recursive: true, force: true });
}
