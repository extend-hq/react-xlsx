// Generates packages/react-xlsx/src/builtin-table-styles.ts from the
// preset table style definitions published in ECMA-376 Part 1 (SpreadsheetML).
//
// Excel workbooks reference built-in table styles by name only, so the viewer
// has to ship the definitions itself. The canonical XML lives in Apache POI
// (Apache-2.0) as presetTableStyles.xml; pass a local path as the first
// argument or let the script download it.
//
//   node scripts/generate-builtin-table-styles.mjs [path/to/presetTableStyles.xml]
import { readFileSync, writeFileSync } from "node:fs";

const SOURCE_URL = "https://raw.githubusercontent.com/apache/poi/trunk/poi-ooxml/src/main/resources/org/apache/poi/xssf/usermodel/presetTableStyles.xml";
const OUTPUT = new URL("../packages/react-xlsx/src/builtin-table-styles.ts", import.meta.url);

async function loadXml() {
  if (process.argv[2]) {
    return readFileSync(process.argv[2], "utf8");
  }
  const response = await fetch(SOURCE_URL);
  if (!response.ok) {
    throw new Error(`Failed to download ${SOURCE_URL}: ${response.status}`);
  }
  return response.text();
}

// The preset file is plain element/attribute XML with no text content, so a
// tiny tag tokenizer is enough to build a tree.
function parseXml(xml) {
  const root = { children: [], name: "#root" };
  const stack = [root];
  const tagPattern = /<\?[^>]*\?>|<!--[\s\S]*?-->|<\/([A-Za-z0-9_:]+)\s*>|<([A-Za-z0-9_:]+)((?:\s+[A-Za-z0-9_:]+="[^"]*")*)\s*(\/?)>/g;
  const attributePattern = /([A-Za-z0-9_:]+)="([^"]*)"/g;
  let match;
  while ((match = tagPattern.exec(xml)) !== null) {
    const [token, closingName, openingName, rawAttributes, selfClosing] = match;
    if (token.startsWith("<?") || token.startsWith("<!--")) {
      continue;
    }
    if (closingName) {
      const node = stack.pop();
      if (!node || node.name !== closingName) {
        throw new Error(`Unbalanced XML near ${token}`);
      }
      continue;
    }
    const attributes = {};
    for (const [, key, value] of rawAttributes.matchAll(attributePattern)) {
      attributes[key] = value;
    }
    const node = { attributes, children: [], name: openingName };
    stack[stack.length - 1].children.push(node);
    if (!selfClosing) {
      stack.push(node);
    }
  }
  if (stack.length !== 1) {
    throw new Error("Unterminated XML element");
  }
  return root;
}

const child = (node, name) => node.children.find((entry) => entry.name === name) ?? null;
const children = (node, name) => node.children.filter((entry) => entry.name === name);

function roundTint(value) {
  return Math.round(Number(value) * 10000) / 10000;
}

// Mirrors parseSpreadsheetColor in images.ts.
function convertColor(node) {
  if (!node) {
    return undefined;
  }
  const color = {};
  if (node.attributes.rgb !== undefined) {
    color.rgb = node.attributes.rgb.toUpperCase();
  }
  if (node.attributes.theme !== undefined) {
    color.theme = Number(node.attributes.theme);
  }
  if (node.attributes.tint !== undefined) {
    color.tint = roundTint(node.attributes.tint);
  }
  if (node.attributes.indexed !== undefined) {
    color.indexed = Number(node.attributes.indexed);
  }
  return Object.keys(color).length > 0 ? color : undefined;
}

// Mirrors parseDifferentialStyle in images.ts for the subset of features the
// presets use (bold/italic fonts, solid fills, and border edges).
function convertDxf(node) {
  const style = {};
  const font = child(node, "font");
  if (font) {
    const nextFont = {};
    if (child(font, "b")) {
      nextFont.bold = true;
    }
    if (child(font, "i")) {
      nextFont.italic = true;
    }
    const color = convertColor(child(font, "color"));
    if (color) {
      nextFont.color = color;
    }
    for (const unsupported of ["sz", "name", "u", "strike"]) {
      if (child(font, unsupported)) {
        throw new Error(`Unsupported font child <${unsupported}> in preset dxf`);
      }
    }
    if (Object.keys(nextFont).length > 0) {
      style.font = nextFont;
    }
  }
  const fill = child(node, "fill");
  if (fill) {
    const patternFill = child(fill, "patternFill");
    if (!patternFill || patternFill.attributes.patternType !== "solid") {
      throw new Error("Unsupported fill in preset dxf");
    }
    const color = convertColor(child(patternFill, "fgColor")) ?? convertColor(child(patternFill, "bgColor"));
    if (color) {
      style.fill = { color, fillType: "solid" };
    }
  }
  const border = child(node, "border");
  if (border) {
    const nextBorder = {};
    for (const edge of ["top", "right", "bottom", "left", "horizontal", "vertical"]) {
      const edgeNode = child(border, edge);
      const edgeStyle = edgeNode?.attributes.style;
      if (!edgeNode || !edgeStyle || edgeStyle === "none") {
        continue;
      }
      nextBorder[edge] = { color: convertColor(child(edgeNode, "color")), style: edgeStyle };
    }
    if (Object.keys(nextBorder).length > 0) {
      style.border = nextBorder;
    }
  }
  if (child(node, "alignment") || child(node, "numFmt")) {
    throw new Error("Unsupported dxf feature in preset");
  }
  return style;
}

function convertPreset(presetNode) {
  const dxfs = children(child(presetNode, "dxfs"), "dxf").map(convertDxf);
  const tableStyle = child(child(presetNode, "tableStyles"), "tableStyle");
  if (!tableStyle || tableStyle.attributes.name !== presetNode.name) {
    throw new Error(`Unexpected tableStyle inside ${presetNode.name}`);
  }
  const elements = {};
  for (const element of children(tableStyle, "tableStyleElement")) {
    // POI's file numbers dxfIds from 1, and its <dxfs count> is not reliable
    // (TableStyleLight14 declares 10 but contains 9), so index the actual list.
    const dxfIndex = Number(element.attributes.dxfId) - 1;
    const dxf = dxfs[dxfIndex];
    if (!dxf) {
      throw new Error(`${presetNode.name}: dxfId ${element.attributes.dxfId} out of range`);
    }
    elements[element.attributes.type] = dxf;
  }
  return elements;
}

function sortKey(name) {
  const match = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(name);
  const familyRank = { Light: 0, Medium: 1, Dark: 2 }[match[1]];
  return familyRank * 100 + Number(match[2]);
}

function formatElement(value) {
  return JSON.stringify(value).replace(/":/g, "\": ").replace(/,"/g, ", \"");
}

const xml = await loadXml();
const root = parseXml(xml);
const presetRoot = child(root, "presetTableStyles");
const presets = presetRoot.children
  .filter((node) => /^TableStyle(Light|Medium|Dark)\d+$/.test(node.name))
  .sort((left, right) => sortKey(left.name) - sortKey(right.name));

if (presets.length !== 60) {
  throw new Error(`Expected 60 table styles, found ${presets.length}`);
}

const lines = [
  "// Generated by scripts/generate-builtin-table-styles.mjs. Do not edit by hand.",
  "//",
  "// Excel's built-in table styles (TableStyleLight1-21, TableStyleMedium1-28 and",
  "// TableStyleDark1-11) as defined in ECMA-376 Part 1, taken from Apache POI's",
  "// presetTableStyles.xml. Workbooks only reference these styles by name, so the",
  "// viewer ships the definitions. Element values use the same shape that",
  "// parseDifferentialStyle produces for custom table styles in styles.xml.",
  "//",
  "// Regenerate with: node scripts/generate-builtin-table-styles.mjs",
  "import type { XlsxTableStyleDefinition } from \"./types\";",
  "",
  "export const BUILTIN_TABLE_STYLES: Record<string, XlsxTableStyleDefinition> = {"
];
presets.forEach((preset, presetIndex) => {
  const elements = convertPreset(preset);
  lines.push(`  ${preset.name}: {`);
  const entries = Object.entries(elements);
  entries.forEach(([type, style], index) => {
    lines.push(`    ${type}: ${formatElement(style)}${index < entries.length - 1 ? "," : ""}`);
  });
  lines.push(`  }${presetIndex < presets.length - 1 ? "," : ""}`);
});
lines.push(
  "};",
  "",
  "/**",
  " * Resolves one of Excel's built-in table styles by name. Workbooks reference",
  " * these by name only, so the definitions ship with the viewer. Returns null for",
  " * any other name, including custom styles defined in the workbook's styles.xml.",
  " */",
  "export function resolveBuiltinTableStyle(name: string | undefined): XlsxTableStyleDefinition | null {",
  "  if (!name || !Object.prototype.hasOwnProperty.call(BUILTIN_TABLE_STYLES, name)) {",
  "    return null;",
  "  }",
  "",
  "  return BUILTIN_TABLE_STYLES[name] ?? null;",
  "}",
  ""
);

writeFileSync(OUTPUT, lines.join("\n"));
console.log(`Wrote ${presets.length} table styles to ${OUTPUT.pathname}`);
