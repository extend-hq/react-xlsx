import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";
import { strToU8, zipSync } from "fflate";
import { initSync, Workbook } from "./duke-sheets/generated/duke_sheets_wasm.js";

initSync({
  module: readFileSync(
    new URL("./duke-sheets/generated/duke_sheets_wasm_bg.wasm", import.meta.url),
  ),
});

function dateWorkbook(date1904: boolean): Uint8Array {
  const parts = {
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    </Types>`,
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    </Relationships>`,
    "xl/workbook.xml": `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <workbookPr date1904="${date1904 ? 1 : 0}"/>
      <sheets><sheet name="Dates" sheetId="1" r:id="rId1"/></sheets>
    </workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    </Relationships>`,
    "xl/styles.xml": `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <numFmts count="3">
        <numFmt numFmtId="164" formatCode="dd mmm yyyy"/>
        <numFmt numFmtId="165" formatCode="yyyy-mm-dd hh:mm:ss.000"/>
        <numFmt numFmtId="166" formatCode="hh:mm:ss.000"/>
      </numFmts>
      <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
      <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
      <borders count="1"><border/></borders>
      <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      <cellXfs count="4">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
        <xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
        <xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
      </cellXfs>
    </styleSheet>`,
    "xl/worksheets/sheet1.xml": `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
      <row r="1"><c r="A1" t="d" s="1"><v>2026-11-01T00:00:00.000Z</v></c></row>
      <row r="2"><c r="A2" t="d" s="1"><v>2026-11-11</v></c></row>
      <row r="3"><c r="A3" t="d" s="2"><v>2026-11-01T12:34:56.125</v></c></row>
      <row r="4"><c r="A4" t="d" s="3"><v>12:34:56.125</v></c></row>
      <row r="5"><c r="A5" t="inlineStr" s="1"><is><t>2026-11-01T00:00:00.000Z</t></is></c></row>
      <row r="6"><c r="A6" t="d" s="1"><v>2026-02-30T00:00:00Z</v></c></row>
      <row r="7"><c r="A7" t="d" s="1"><f>DATE(2026,11,1)</f><v>2026-11-01T00:00:00Z</v></c></row>
      <row r="8"><c r="A8" t="n" s="1"><v>${date1904 ? 44865 : 46327}</v></c></row>
      <row r="9"><c r="A9" t="d"><v>1900-01-01</v></c></row>
      <row r="10"><c r="A10" t="d"><v>1900-02-28</v></c></row>
      <row r="11"><c r="A11" t="d"><v>1900-03-01</v></c></row>
      <row r="12"><c r="A12" t="d"><v>1904-01-01</v></c></row>
      <row r="13"><c r="A13" t="d"><v>1900-02-29</v></c></row>
      <row r="14"><c r="A14" t="d"><v>23:59:60</v></c></row>
      <row r="15"><c r="A15" t="d"><v>2026-11-01T00:00:00+02:00</v></c></row>
    </sheetData></worksheet>`,
  };
  return zipSync(
    Object.fromEntries(Object.entries(parts).map(([name, xml]) => [name, strToU8(xml)])),
  );
}

for (const date1904 of [false, true]) {
  for (const [mode, load] of [
    ["viewer", Workbook.fromBytesForViewer],
    ["editable", Workbook.fromBytes],
    [
      "snapshot",
      (bytes: Uint8Array) => {
        const source = Workbook.fromBytesForViewer(bytes);
        try {
          return Workbook.fromSnapshot(source.exportSnapshot());
        } finally {
          source.free();
        }
      },
    ],
  ] as const) {
    test(`${mode} imports typed ISO dates using the ${date1904 ? 1904 : 1900} date system`, () => {
      const workbook = load(dateWorkbook(date1904));
      const sheet = workbook.getSheet(0);
      try {
        const serial = date1904 ? 44865 : 46327;
        const expected = [serial, serial + 10, serial + 45296.125 / 86400, 45296.125 / 86400];
        for (const [row, value] of expected.entries()) {
          const cell = sheet.getCellAt(row, 0);
          try {
            assert.ok(cell.is_number, `A${row + 1} should be numeric, got ${cell.toString()}`);
            assert.ok(Math.abs(cell.asNumber()! - value) < 1e-9);
          } finally {
            cell.free();
          }
        }
        assert.equal(sheet.getFormattedValue("A1"), "01 Nov 2026");
        assert.equal(sheet.getFormattedValue("A2"), "11 Nov 2026");
        assert.equal(sheet.getFormattedValue("A3"), "2026-11-01 12:34:56.125");
        assert.equal(sheet.getFormattedValue("A4"), "12:34:56.125");
        assert.equal(sheet.getFormattedValue("A5"), "2026-11-01T00:00:00.000Z");
        assert.equal(sheet.getFormattedValue("A6"), "2026-02-30T00:00:00Z");
        assert.equal(sheet.getFormattedValue("A7"), "01 Nov 2026");
        assert.equal(sheet.getFormattedValue("A8"), "01 Nov 2026");
        assert.equal(sheet.getFormattedValue("A12"), date1904 ? "0" : "1462");
        assert.equal(sheet.getFormattedValue("A13"), "1900-02-29");
        assert.equal(sheet.getFormattedValue("A14"), "23:59:60");
        // Unsupported offsets retain their text instead of silently shifting the time.
        assert.equal(sheet.getFormattedValue("A15"), "2026-11-01T00:00:00+02:00");
        if (!date1904) {
          assert.equal(sheet.getFormattedValue("A9"), "1");
          assert.equal(sheet.getFormattedValue("A10"), "59");
          assert.equal(sheet.getFormattedValue("A11"), "61");
        }
      } finally {
        sheet.free();
        workbook.free();
      }
    });
  }
}
