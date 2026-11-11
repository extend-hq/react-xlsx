import { spawnSync } from 'node:child_process';
import { copyFileSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir, homedir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { dukeSheetsSourceDigest, patchNames, sha256, upstream } from './duke-sheets-provenance.mjs';
const root = fileURLToPath(new URL('../', import.meta.url));
const output = join(root, 'packages/react-xlsx/src/duke-sheets/generated');
const checkout = mkdtempSync(join(tmpdir(), 'react-xlsx-duke-'));
const run = (command, args, cwd = checkout) => {
  const result = spawnSync(command, args, { cwd, stdio: 'inherit', env: { ...process.env, CARGO_ENCODED_RUSTFLAGS: `--remap-path-prefix=${homedir()}=/home/builder\x1f--remap-path-prefix=${checkout}=.` } });
  if (result.error || result.status !== 0) throw result.error ?? new Error(`${command} exited ${result.status}`);
};
try {
  const sourceDigest = dukeSheetsSourceDigest();
  run('git', ['init', '--quiet']);
  run('git', ['fetch', '--depth=1', upstream.repository, upstream.commit]);
  run('git', ['checkout', '--detach', 'FETCH_HEAD']);
  for (const name of patchNames()) {
    run('git', ['apply', '--check', join(root, 'patches/duke-sheets', name)]);
    run('git', ['apply', join(root, 'patches/duke-sheets', name)]);
  }
  mkdirSync(output, { recursive: true });
  run('wasm-pack', ['build', 'bindings/wasm', '--target', 'web', '--release', '--no-pack', '--out-dir', output, '--', '--locked']);
  for (const name of ['.gitignore', 'package.json']) rmSync(join(output, name), { force: true });
  copyFileSync(join(checkout, 'LICENSE'), join(output, 'LICENSE'));
  const names = ['duke_sheets_wasm.js', 'duke_sheets_wasm.d.ts', 'duke_sheets_wasm_bg.wasm', 'duke_sheets_wasm_bg.wasm.d.ts', 'LICENSE'];
  if (sourceDigest !== dukeSheetsSourceDigest()) throw new Error('Parser patches changed during compilation.');
  writeFileSync(join(output, 'provenance.json'), JSON.stringify({ ...upstream, sourceDigest, outputs: Object.fromEntries(names.map(name => [name, sha256(readFileSync(join(output, name)))])) }, null, 2) + '\n');
} finally { rmSync(checkout, { recursive: true, force: true }); }
