import { createHash } from 'node:crypto';
import { readFileSync, readdirSync } from 'node:fs';
export const upstream = { repository: 'https://github.com/guseggert/duke-sheets.git', commit: 'fec1cd04dbefa5eb65a41d29399ca4a7edc220f2', version: '0.1.24' };
export const patchNames = () => readdirSync(new URL('../patches/duke-sheets/', import.meta.url)).filter(name => name.endsWith('.patch')).sort();
export const sha256 = bytes => createHash('sha256').update(bytes).digest('hex');
export function dukeSheetsSourceDigest() {
  const paths = ['scripts/build-duke-sheets.mjs', 'scripts/duke-sheets-provenance.mjs', ...patchNames().map(name => `patches/duke-sheets/${name}`)];
  return sha256(Buffer.concat(paths.flatMap(path => [Buffer.from(`${path}\0`), readFileSync(new URL(`../${path}`, import.meta.url))])));
}
