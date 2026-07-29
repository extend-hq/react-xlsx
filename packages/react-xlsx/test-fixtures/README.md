# Test fixtures

Workbooks kept for manual verification of reader edge cases. Both files here are
valid OPC packages whose ZIP entry names use the Windows path separator
(`xl\worksheets\sheet1.xml`) instead of the forward slash required by the ZIP
specification (APPNOTE 4.4.17.1) and OPC.

Before the fix in `src/zip-entry-names.ts`, loading either file failed with:

```
Failed to read file: Missing required part: xl/worksheets/sheet1.xml
```

| File | Size | Notes |
| --- | --- | --- |
| `backslash-entry-names-minimal.xlsx` | 1.6 KB | Synthetic minimal reproduction: 5 parts, 1 sheet, no styles or shared strings. |
| `backslash-entry-names-reported.xlsx` | 6.7 KB | The workbook from the original report. 3 sheets with non-ASCII names, shared strings and styles. |

To verify, drop either file into the playground (`pnpm dev`).

Inspecting the entry names directly:

```js
import { unzipSync } from "fflate";
Object.keys(unzipSync(new Uint8Array(await file.arrayBuffer())));
// [ '[Content_Types].xml', '_rels\\.rels', 'xl\\workbook.xml', ... ]
```
