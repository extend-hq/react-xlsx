import { readFileSync } from 'node:fs';
import { dukeSheetsSourceDigest, sha256, upstream } from './duke-sheets-provenance.mjs';
const root = new URL('../packages/react-xlsx/src/duke-sheets/generated/', import.meta.url);
const provenance = JSON.parse(readFileSync(new URL('provenance.json', root), 'utf8'));
if (provenance.commit !== upstream.commit || provenance.sourceDigest !== dukeSheetsSourceDigest()) throw new Error('Spreadsheet parser bindings need rebuilding with pnpm build:duke-sheets.');
for (const [name, digest] of Object.entries(provenance.outputs)) {
  if (sha256(readFileSync(new URL(name, root))) !== digest) throw new Error('Spreadsheet parser output failed its integrity check.');
}
