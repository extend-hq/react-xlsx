import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { before, describe, it } from "node:test";
import initSheetsWasm, { Workbook } from "@dukelib/sheets-wasm";
import { strToU8, zipSync } from "fflate";
import { safeCalculate } from "./safe-calculate.ts";

const CURRENCY_STYLE = 1;

function buildWorkbook({ formula, cachedValue, definedNames = "" }: {
  formula: string;
  cachedValue?: string;
  definedNames?: string;
}): Uint8Array {
  const cached = cachedValue === undefined ? "" : `<v>${cachedValue}</v>`;
  return zipSync({
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
    ),
    "xl/workbook.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
${definedNames ? `<definedNames>${definedNames}</definedNames>` : ""}
</workbook>`
    ),
    "xl/_rels/workbook.xml.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
    ),
    "xl/styles.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0"/></numFmts>
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>
</styleSheet>`
    ),
    "xl/worksheets/sheet1.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1"><v>2500</v></c><c r="B1" s="${CURRENCY_STYLE}"><f>${formula}</f>${cached}</c></row>
<row r="2"><c r="A2"><v>1000</v></c></row>
<row r="3"><c r="A3"><v>250</v></c></row>
</sheetData>
</worksheet>`
    )
  });
}

function calculate(bytes: Uint8Array) {
  return safeCalculate(Workbook.fromBytes(bytes), { reparse: () => Workbook.fromBytes(bytes) });
}

describe("safeCalculate", () => {
  before(async () => {
    const wasmFile = readFileSync(
      new URL(import.meta.resolve("@dukelib/sheets-wasm/duke_sheets_wasm_bg.wasm"))
    );
    await initSheetsWasm({ module_or_path: wasmFile });
  });

  it("keeps saved formula values when recalculation turns them into errors", () => {
    const result = calculate(
      buildWorkbook({
        formula: "TotalIncome",
        cachedValue: "3750",
        definedNames: `<definedName name="TotalIncome">SUM(Sheet1!$A$1:$A$3)</definedName>`
      })
    );

    assert.equal(result.calculated, false);
    assert.equal(result.skipReason, "calculate-regressed");
    assert.equal(result.workbook.getSheet(0).getFormattedValueAt(0, 1), "$3,750");
  });

  it("still calculates formulas that have no saved value", () => {
    const result = calculate(buildWorkbook({ formula: "SUM(A1:A3)" }));

    assert.equal(result.calculated, true);
    assert.equal(result.skipReason, null);
    assert.equal(result.workbook.getSheet(0).getFormattedValueAt(0, 1), "$3,750");
  });

  it("prefers the recalculated value when calculation succeeds", () => {
    const result = calculate(buildWorkbook({ formula: "SUM(A1:A3)", cachedValue: "1" }));

    assert.equal(result.calculated, true);
    assert.equal(result.workbook.getSheet(0).getFormattedValueAt(0, 1), "$3,750");
  });
});
