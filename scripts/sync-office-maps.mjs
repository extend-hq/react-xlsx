import { createHash } from 'node:crypto';
import { copyFileSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { pathToFileURL, fileURLToPath } from 'node:url';

const root = resolve(process.argv[2] ?? '../extend-office');
const destination = fileURLToPath(new URL('../packages/react-xlsx/src/office-maps/', import.meta.url));
const { sourceDigest } = await import(pathToFileURL(resolve(root, 'browser/source-digest.mjs')));
const sourceSha256 = sourceDigest();
if (sourceSha256 !== readFileSync(resolve(root, 'browser/generated/sources.sha256'), 'utf8').trim()) {
  throw Error('Build the Office Charts browser artifacts before syncing.');
}
mkdirSync(destination, { recursive: true });
for (const name of ['office_charts.js', 'office_charts_bg.wasm', 'payload.js']) {
  copyFileSync(resolve(root, 'browser/generated/full', name), resolve(destination, name));
}
for (const name of ['LICENSE', 'MAP-DATA-LICENSE']) {
  copyFileSync(resolve(root, name), resolve(destination, name));
}
copyFileSync(resolve(root, 'browser/loader.js'), resolve(destination, 'loader.js'));
writeFileSync(resolve(destination, 'index.js'), `import init, { renderChartSvg } from './office_charts.js';\nimport { createLoader } from './loader.js';\nexport const loadOfficeCharts = createLoader(init, renderChartSvg, () => import('./payload.js'));\n`);
writeFileSync(resolve(destination, 'index.d.ts'), `export { loadOfficeCharts } from '@extend-ai/extend-office-charts';\n`);
const binary = readFileSync(resolve(destination, 'office_charts_bg.wasm'));
writeFileSync(resolve(destination, 'provenance.json'), JSON.stringify({
  package: '@extend-ai/extend-office-charts',
  version: JSON.parse(readFileSync(resolve(root, 'package.json'), 'utf8')).version,
  sourceSha256,
  wasmSha256: createHash('sha256').update(binary).digest('hex'),
  wasmBytes: binary.length,
  features: ['map-fallback'],
}, null, 2) + '\n');
console.log(`Synced Rust map renderer (${binary.length} bytes).`);
