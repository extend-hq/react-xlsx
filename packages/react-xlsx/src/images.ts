import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { resolveWorkbookColor } from "./colors";
import type {
  XlsxCellAddress,
  XlsxConditionalDataBarRule,
  XlsxConditionalFormatIcon,
  XlsxConditionalFormatRule,
  XlsxConditionalFormatValueObject,
  XlsxConditionalIconSetRule,
  XlsxCellRange,
  XlsxFormControl,
  XlsxImage,
  XlsxImageRect,
  XlsxImageResizeHandlePosition,
  XlsxResolvedCellStyle,
  XlsxShape,
  XlsxSparkline,
  XlsxTableStyleDefinition,
  XlsxThemePalette
} from "./types";

const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const DRAWING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const VML_DRAWING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
const CTRL_PROP_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp";
const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const HYPERLINK_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const DRAWING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawing+xml";
const EMU_PER_PIXEL = 9525;
const MIN_COL_WIDTH_PX = 30;
const MIN_ROW_HEIGHT_PX = 16;
const DEFAULT_COL_WIDTH_EMU = 64 * EMU_PER_PIXEL;
const DEFAULT_ROW_HEIGHT_EMU = 20 * EMU_PER_PIXEL;
const DEFAULT_COLUMN_CHARACTER_WIDTH_PX = 7;
const columnCharacterWidthCache = new Map<string, number>();

function resolveDeviceGridlineThicknessPx() {
  if (typeof window === "undefined") {
    return 1;
  }

  const devicePixelRatio = window.devicePixelRatio;
  if (!Number.isFinite(devicePixelRatio) || devicePixelRatio <= 0) {
    return 1;
  }

  return 1 / devicePixelRatio;
}

function measureColumnCharacterWidthPx(fontFamily?: string | null, fontSizePt?: number | null) {
  const normalizedFamily = typeof fontFamily === "string" && fontFamily.trim().length > 0
    ? fontFamily.trim()
    : "Calibri";
  const normalizedSizePt = typeof fontSizePt === "number" && Number.isFinite(fontSizePt) && fontSizePt > 0
    ? fontSizePt
    : 11;
  const cacheKey = `${normalizedFamily}|${normalizedSizePt}`;
  const cached = columnCharacterWidthCache.get(cacheKey);
  if (cached !== undefined) {
    return cached;
  }

  const fontSizePx = normalizedSizePt * (96 / 72);
  const font = `${fontSizePx}px "${normalizedFamily}"`;
  let width = DEFAULT_COLUMN_CHARACTER_WIDTH_PX;

  try {
    const context = typeof document !== "undefined"
      ? document.createElement("canvas").getContext("2d")
      : typeof OffscreenCanvas !== "undefined"
        ? new OffscreenCanvas(32, 32).getContext("2d")
        : null;
    if (context) {
      context.font = font;
      width = Math.max(1, context.measureText("0").width);
    }
  } catch {
    width = DEFAULT_COLUMN_CHARACTER_WIDTH_PX;
  }

  columnCharacterWidthCache.set(cacheKey, width);
  return width;
}

function sheetColumnWidthToPixels(width: number, columnCharacterWidthPx = DEFAULT_COLUMN_CHARACTER_WIDTH_PX) {
  if (!Number.isFinite(width) || width <= 0) {
    return MIN_COL_WIDTH_PX;
  }

  const digitWidth = Math.max(1, columnCharacterWidthPx);
  const pixels = width < 1
    ? Math.floor(width * (digitWidth + 5) + 0.5)
    : Math.floor(((256 * width + Math.floor(128 / digitWidth)) / 256) * digitWidth);
  return Math.max(MIN_COL_WIDTH_PX, pixels);
}

type ArchiveEntries = Record<string, Uint8Array>;

type ContentTypesState = {
  defaultEntries: Map<string, string>;
  overrideEntries: Map<string, string>;
};

type RelationshipRecord = {
  id: string;
  target: string;
  targetMode: string | null;
  type: string;
};

type WorkbookSheetInfo = {
  name: string;
  path: string;
};

type WorkbookSheetState = {
  cachedFormulaValues: Record<string, string>;
  columnWidthCharacterWidthPx?: number;
  colWidthOverridesPx: Record<number, number>;
  colStyleIds: Record<number, number>;
  conditionalFormatRules: XlsxConditionalFormatRule[];
  defaultColWidthPx: number;
  defaultRowHeightPx: number;
  hasHorizontalMerges: boolean;
  hasVerticalMerges: boolean;
  maxHorizontalMergeEndCol: number;
  maxVerticalMergeEndRow: number;
  hiddenCols: number[];
  hiddenRows: number[];
  rowHeightOverridesPx: Record<number, number>;
  rowStyleIds: Record<number, number>;
  showGridLines: boolean;
  sparklines: XlsxSparkline[];
  zoomScale: number;
};

type ParseWorkbookStructureOptions = {
  includeCachedFormulaValues?: boolean;
  themePalette?: XlsxThemePalette | null;
};

type ThemeState = {
  colors: Map<string, string>;
  majorLatinFont: string | null;
  minorLatinFont: string | null;
};

type DrawingColor = {
  color: string;
  opacity: number;
};

type DrawingRectEmu = {
  cx: number;
  cy: number;
  x: number;
  y: number;
};

type ShapeVectorPath = {
  path: string;
  viewBox: {
    height: number;
    width: number;
  };
};

type GroupTransform = {
  chCx: number;
  chCy: number;
  chX: number;
  chY: number;
  cx: number;
  cy: number;
  scaleX: number;
  scaleY: number;
  x: number;
  y: number;
};

type XlsxImageAttachment = {
  drawingPath: string;
  drawingRelsPath: string | null;
  mediaPaths: string[];
};

type WorkbookImageOrigin = {
  anchorIndex: number;
  workbookSheetIndex: number;
};

type ParsedSheetFormControl = {
  anchor: XlsxFormControl["anchor"] | null;
  controlRelationshipId: string | null;
  name?: string;
  shapeId: number | null;
};

type ParsedCtrlProp = {
  checked?: boolean;
  linkedCell?: string;
  objectType?: string;
};

type ParsedVmlFormControl = {
  checked?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  hidden: boolean;
  label?: string;
  linkedCell?: string;
  objectType?: string;
  shapeId: number | null;
  textAlign?: "center" | "left" | "right";
  textColor?: string;
  zIndex: number;
};

export type WorkbookImageSheetOrigin = {
  attachments: XlsxImageAttachment[];
  workbookSheetIndex: number;
};

export type WorkbookTableMetadata = {
  displayName?: string;
  headerRowCellStyle?: string;
  name?: string;
  reference?: string;
};

export type WorkbookImageAssets = {
  archive: ArchiveEntries;
  formControlsByWorkbookSheetIndex: XlsxFormControl[][];
  imageOriginsById: Map<string, WorkbookImageOrigin>;
  imagesByWorkbookSheetIndex: XlsxImage[][];
  namedCellStyleByName: Record<string, XlsxResolvedCellStyle>;
  objectUrls: string[];
  shapesByWorkbookSheetIndex: XlsxShape[][];
  sheetStatesByWorkbookSheetIndex: Array<WorkbookSheetState | null>;
  sheetOrigins: Array<WorkbookImageSheetOrigin | null>;
  styleById: Record<number, XlsxResolvedCellStyle>;
  tableMetadataByWorkbookSheetIndex: WorkbookTableMetadata[][];
  tableStyleByName: Record<string, XlsxTableStyleDefinition>;
  themePalette: XlsxThemePalette;
};

export type WorkbookStructureAssets = Pick<
  WorkbookImageAssets,
  | "namedCellStyleByName"
  | "sheetStatesByWorkbookSheetIndex"
  | "styleById"
  | "tableMetadataByWorkbookSheetIndex"
  | "tableStyleByName"
  | "themePalette"
>;

export type WorkbookChartStyleAssets = Pick<
  WorkbookImageAssets,
  | "archive"
  | "sheetOrigins"
  | "themePalette"
>;

function buildThemePalette(theme: ThemeState): XlsxThemePalette {
  const themeOrder = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  const colorsByIndex: Record<number, string> = {};

  themeOrder.forEach((key, index) => {
    const color = theme.colors.get(key);
    if (color) {
      colorsByIndex[index] = color;
    }
  });

  return {
    colorsByIndex,
    majorLatinFont: theme.majorLatinFont ?? undefined,
    minorLatinFont: theme.minorLatinFont ?? undefined
  };
}

function cloneBytes(bytes: Uint8Array) {
  const nextBytes = new Uint8Array(bytes.byteLength);
  nextBytes.set(bytes);
  return nextBytes;
}

function normalizeArchivePath(path: string) {
  return path.replace(/\\/g, "/").replace(/^\/+/, "");
}

function joinArchivePath(...parts: string[]) {
  return normalizeArchivePath(parts.join("/"));
}

function dirname(path: string) {
  const normalized = normalizeArchivePath(path);
  const lastSlash = normalized.lastIndexOf("/");
  return lastSlash >= 0 ? normalized.slice(0, lastSlash) : "";
}

function resolveArchiveTarget(baseDocumentPath: string, target: string) {
  if (!target) {
    return normalizeArchivePath(baseDocumentPath);
  }

  if (target.startsWith("#")) {
    return target;
  }

  if (target.startsWith("/")) {
    return normalizeArchivePath(target);
  }

  const baseParts = dirname(baseDocumentPath).split("/").filter(Boolean);
  for (const segment of target.split("/")) {
    if (!segment || segment === ".") {
      continue;
    }
    if (segment === "..") {
      baseParts.pop();
      continue;
    }
    baseParts.push(segment);
  }

  return normalizeArchivePath(baseParts.join("/"));
}

function relativeArchivePath(fromDocumentPath: string, toPath: string) {
  const fromParts = dirname(fromDocumentPath).split("/").filter(Boolean);
  const toParts = normalizeArchivePath(toPath).split("/").filter(Boolean);
  let shared = 0;
  while (shared < fromParts.length && shared < toParts.length && fromParts[shared] === toParts[shared]) {
    shared += 1;
  }

  const upSegments = fromParts.slice(shared).map(() => "..");
  const downSegments = toParts.slice(shared);
  return [...upSegments, ...downSegments].join("/") || ".";
}

function relsPathForDocument(documentPath: string) {
  const baseName = documentPath.split("/").pop();
  const parentDir = dirname(documentPath);
  return joinArchivePath(parentDir, "_rels", `${baseName}.rels`);
}

function parseXml(xml: string) {
  const parser = new DOMParser();
  const document = parser.parseFromString(xml, "application/xml");
  if (document.querySelector("parsererror")) {
    return null;
  }
  return document;
}

function serializeXml(document: XMLDocument) {
  return new XMLSerializer().serializeToString(document);
}

function readArchiveText(archive: ArchiveEntries, path: string) {
  const entry = archive[path];
  return entry ? strFromU8(entry) : null;
}

function parseColumnReference(reference: string) {
  let value = 0;
  for (const character of reference.toUpperCase()) {
    if (character < "A" || character > "Z") {
      return null;
    }
    value = value * 26 + (character.charCodeAt(0) - 64);
  }
  return value > 0 ? value - 1 : null;
}

function parseA1CellReference(reference: string) {
  const match = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(reference.trim());
  if (!match) {
    return null;
  }

  const col = parseColumnReference(match[1] ?? "");
  const row = Number(match[2] ?? Number.NaN) - 1;
  if (col === null || !Number.isFinite(row) || row < 0) {
    return null;
  }

  return { col, row };
}

function parseA1RangeReference(reference: string) {
  const [startRef, endRef] = reference.split(":");
  const start = parseA1CellReference(startRef ?? "");
  const end = parseA1CellReference(endRef ?? startRef ?? "");
  return start && end ? { end, start } : null;
}

function stripSheetNameFromFormulaReference(reference: string) {
  const trimmed = reference.trim();
  const bangIndex = trimmed.lastIndexOf("!");
  return bangIndex >= 0 ? trimmed.slice(bangIndex + 1) : trimmed;
}

function parseFormulaCellReference(reference: string) {
  const normalized = stripSheetNameFromFormulaReference(reference).split(/\s+/)[0] ?? "";
  return parseA1CellReference(normalized);
}

function parseFormulaRangeReference(reference: string) {
  return parseA1RangeReference(stripSheetNameFromFormulaReference(reference));
}

function isElementNode(node: Node | null | undefined): node is Element {
  return Boolean(node && node.nodeType === 1);
}

function getLocalElements(parent: Document | Element, localName: string) {
  return Array.from(parent.getElementsByTagName("*")).filter((node): node is Element => isElementNode(node) && node.localName === localName);
}

function getChildElements(parent: Element, localName: string) {
  return Array.from(parent.childNodes).filter((node): node is Element => isElementNode(node) && node.localName === localName);
}

function getFirstChild(parent: Element, localName: string) {
  return getChildElements(parent, localName)[0] ?? null;
}

function getFirstDescendant(parent: Document | Element, localName: string) {
  return getLocalElements(parent, localName)[0] ?? null;
}

function readFeaturePropertyBagCheckboxComplements(archive: ArchiveEntries) {
  const xml = readArchiveText(archive, "xl/featurePropertyBag/featurePropertyBag.xml");
  if (!xml) {
    return new Set<number>();
  }

  const document = parseXml(xml);
  if (!document?.documentElement) {
    return new Set<number>();
  }

  const bagNodes = getChildElements(document.documentElement, "bag");
  const bagTypeById = bagNodes.map((node) => node.getAttribute("type") ?? "");
  const checkboxComplementIndices = new Set<number>();

  const xfComplementsBag = bagNodes.find((node) => node.getAttribute("type") === "XFComplements") ?? null;
  const mappedBagIds = xfComplementsBag
    ? getLocalElements(xfComplementsBag, "bagId")
      .map((node) => Number(node.textContent ?? Number.NaN))
      .filter((value) => Number.isFinite(value))
    : [];

  mappedBagIds.forEach((bagId, complementIndex) => {
    const xfComplementBag = bagNodes[bagId];
    if (!xfComplementBag || bagTypeById[bagId] !== "XFComplement") {
      return;
    }

    const xfControlsBagId = getLocalElements(xfComplementBag, "bagId")
      .map((node) => Number(node.textContent ?? Number.NaN))
      .find((value) => Number.isFinite(value));
    if (xfControlsBagId === undefined) {
      return;
    }

    const xfControlsBag = bagNodes[xfControlsBagId];
    if (!xfControlsBag || bagTypeById[xfControlsBagId] !== "XFControls") {
      return;
    }

    const cellControlBagId = getLocalElements(xfControlsBag, "bagId")
      .map((node) => Number(node.textContent ?? Number.NaN))
      .find((value) => Number.isFinite(value));
    if (cellControlBagId === undefined) {
      return;
    }

    if (bagTypeById[cellControlBagId] === "Checkbox") {
      checkboxComplementIndices.add(complementIndex);
    }
  });

  return checkboxComplementIndices;
}

function getRelationshipId(element: Element) {
  return element.getAttributeNS(REL_NS, "id") ?? element.getAttribute("r:id") ?? element.getAttribute("id");
}

function getEmbeddedRelationshipId(element: Element) {
  return element.getAttributeNS(REL_NS, "embed") ?? element.getAttribute("r:embed") ?? element.getAttribute("embed");
}

function setChildText(parent: Element, localName: string, value: string) {
  const child = getFirstChild(parent, localName);
  if (child) {
    child.textContent = value;
  }
}

function updateMarkerElement(element: Element | null, marker: { col: number; colOffsetEmu: number; row: number; rowOffsetEmu: number }) {
  if (!element) {
    return;
  }

  setChildText(element, "col", String(Math.max(0, marker.col)));
  setChildText(element, "colOff", String(Math.max(0, Math.round(marker.colOffsetEmu))));
  setChildText(element, "row", String(Math.max(0, marker.row)));
  setChildText(element, "rowOff", String(Math.max(0, Math.round(marker.rowOffsetEmu))));
}

function parseContentTypes(archive: ArchiveEntries): ContentTypesState {
  const xml = readArchiveText(archive, "[Content_Types].xml");
  const defaultEntries = new Map<string, string>();
  const overrideEntries = new Map<string, string>();
  if (!xml) {
    return { defaultEntries, overrideEntries };
  }

  const document = parseXml(xml);
  if (!document) {
    return { defaultEntries, overrideEntries };
  }

  for (const defaultNode of getLocalElements(document, "Default")) {
    const extension = defaultNode.getAttribute("Extension");
    const contentType = defaultNode.getAttribute("ContentType");
    if (extension && contentType) {
      defaultEntries.set(extension.toLowerCase(), contentType);
    }
  }

  for (const overrideNode of getLocalElements(document, "Override")) {
    const partName = overrideNode.getAttribute("PartName");
    const contentType = overrideNode.getAttribute("ContentType");
    if (partName && contentType) {
      overrideEntries.set(normalizeArchivePath(partName), contentType);
    }
  }

  return { defaultEntries, overrideEntries };
}

function resolveContentType(contentTypes: ContentTypesState, path: string) {
  const normalized = normalizeArchivePath(path);
  const override = contentTypes.overrideEntries.get(normalized);
  if (override) {
    return override;
  }

  const extension = normalized.split(".").pop()?.toLowerCase();
  if (!extension) {
    return "application/octet-stream";
  }

  return contentTypes.defaultEntries.get(extension) ?? "application/octet-stream";
}

function parseRelationships(archive: ArchiveEntries, relsPath: string, baseDocumentPath: string) {
  const xml = readArchiveText(archive, relsPath);
  const relationships = new Map<string, RelationshipRecord>();
  if (!xml) {
    return relationships;
  }

  const document = parseXml(xml);
  if (!document) {
    return relationships;
  }

  for (const relationshipNode of getLocalElements(document, "Relationship")) {
    const id = relationshipNode.getAttribute("Id");
    const target = relationshipNode.getAttribute("Target");
    const type = relationshipNode.getAttribute("Type");
    if (!id || !target || !type) {
      continue;
    }

    relationships.set(id, {
      id,
      target: resolveArchiveTarget(baseDocumentPath, target),
      targetMode: relationshipNode.getAttribute("TargetMode"),
      type
    });
  }

  return relationships;
}

function parseWorkbookSheets(archive: ArchiveEntries) {
  const workbookXml = readArchiveText(archive, "xl/workbook.xml");
  if (!workbookXml) {
    return [];
  }

  const workbookDocument = parseXml(workbookXml);
  if (!workbookDocument) {
    return [];
  }

  const workbookRelationships = parseRelationships(archive, "xl/_rels/workbook.xml.rels", "xl/workbook.xml");
  const sheets: WorkbookSheetInfo[] = [];

  for (const sheetNode of getLocalElements(workbookDocument, "sheet")) {
    const relationshipId = getRelationshipId(sheetNode);
    if (!relationshipId) {
      continue;
    }

    const relationship = workbookRelationships.get(relationshipId);
    if (!relationship) {
      continue;
    }

    sheets.push({
      name: sheetNode.getAttribute("name") ?? `Sheet ${sheets.length + 1}`,
      path: relationship.target
    });
  }

  return sheets;
}

function parseWorkbookTheme(archive: ArchiveEntries): ThemeState {
  const defaultTheme: ThemeState = {
    colors: new Map([
      ["accent1", "#5b9bd5"],
      ["accent2", "#ed7d31"],
      ["accent3", "#a5a5a5"],
      ["accent4", "#ffc000"],
      ["accent5", "#4472c4"],
      ["accent6", "#70ad47"],
      ["bg1", "#ffffff"],
      ["bg2", "#e7e6e6"],
      ["dk1", "#000000"],
      ["dk2", "#6e747a"],
      ["folHlink", "#993366"],
      ["hlink", "#085296"],
      ["lt1", "#ffffff"],
      ["lt2", "#e7e6e6"],
      ["tx1", "#000000"],
      ["tx2", "#6e747a"]
    ]),
    majorLatinFont: null,
    minorLatinFont: null
  };

  const themeXml = readArchiveText(archive, "xl/theme/theme1.xml");
  if (!themeXml) {
    return defaultTheme;
  }

  const themeDocument = parseXml(themeXml);
  if (!themeDocument) {
    return defaultTheme;
  }

  const colors = new Map(defaultTheme.colors);
  const colorSchemeNode = getLocalElements(themeDocument, "clrScheme")[0] ?? null;
  if (colorSchemeNode) {
    for (const colorNode of Array.from(colorSchemeNode.childNodes).filter(isElementNode)) {
      const key = colorNode.localName;
      const srgbNode = getFirstChild(colorNode, "srgbClr");
      const sysNode = getFirstChild(colorNode, "sysClr");
      const hex = srgbNode?.getAttribute("val") ?? sysNode?.getAttribute("lastClr");
      if (hex) {
        colors.set(key, normalizeHexColor(hex));
      }
    }
  }

  const fontSchemeNode = getLocalElements(themeDocument, "fontScheme")[0] ?? null;
  const majorLatinFont = getFirstChild(getFirstChild(fontSchemeNode, "majorFont"), "latin")?.getAttribute("typeface") ?? null;
  const minorLatinFont = getFirstChild(getFirstChild(fontSchemeNode, "minorFont"), "latin")?.getAttribute("typeface") ?? null;

  colors.set("bg1", colors.get("lt1") ?? defaultTheme.colors.get("bg1") ?? "#ffffff");
  colors.set("tx1", colors.get("dk1") ?? defaultTheme.colors.get("tx1") ?? "#000000");
  colors.set("bg2", colors.get("lt2") ?? defaultTheme.colors.get("bg2") ?? "#e7e6e6");
  colors.set("tx2", colors.get("dk2") ?? defaultTheme.colors.get("tx2") ?? "#6e747a");

  return {
    colors,
    majorLatinFont,
    minorLatinFont
  };
}

function parseSpreadsheetColor(node: Element | null) {
  if (!node) {
    return undefined;
  }

  const color: Record<string, unknown> = {};
  const rgb = node.getAttribute("rgb");
  const theme = node.getAttribute("theme");
  const tint = node.getAttribute("tint");
  const indexed = node.getAttribute("indexed");
  if (rgb) {
    color.rgb = normalizeHexColor(rgb);
  }
  if (theme !== null) {
    color.theme = Number(theme);
  }
  if (tint !== null) {
    color.tint = Number(tint);
  }
  if (indexed !== null) {
    color.indexed = Number(indexed);
  }

  return Object.keys(color).length > 0 ? color : undefined;
}

function hasEnabledSpreadsheetFlag(node: Element | null) {
  if (!node) {
    return false;
  }

  const value = node.getAttribute("val");
  return value === null || (value !== "0" && value !== "false");
}

function parseSheetSparklines(
  document: XMLDocument,
  themePalette?: XlsxThemePalette | null
) {
  const sparklines: XlsxSparkline[] = [];

  for (const groupNode of getLocalElements(document, "sparklineGroup")) {
    const rawType = groupNode.getAttribute("type");
    const sparklineType: XlsxSparkline["type"] = rawType === "column"
      ? "column"
      : rawType === "stacked"
        ? "winLoss"
        : "line";

    const markersNode = getFirstChild(groupNode, "markers");
    const negativeNode = getFirstChild(groupNode, "negative");
    const colorSeries = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorSeries")), themePalette);
    const colorNegative = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorNegative")), themePalette);
    const colorMarkers = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorMarkers")), themePalette);
    const colorFirst = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorFirst")), themePalette);
    const colorLast = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorLast")), themePalette);
    const colorHigh = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorHigh")), themePalette);
    const colorLow = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorLow")), themePalette);
    const sparklineCollectionNode = getFirstChild(groupNode, "sparklines");
    if (!sparklineCollectionNode) {
      continue;
    }

    for (const sparklineNode of getChildElements(sparklineCollectionNode, "sparkline")) {
      const formula = getFirstChild(sparklineNode, "f")?.textContent ?? "";
      const targetReference = getFirstChild(sparklineNode, "sqref")?.textContent ?? "";
      const range = parseFormulaRangeReference(formula);
      const target = parseFormulaCellReference(targetReference);
      if (!range || !target) {
        continue;
      }

      sparklines.push({
        color: colorSeries ?? undefined,
        firstColor: colorFirst ?? undefined,
        highColor: colorHigh ?? undefined,
        lastColor: colorLast ?? undefined,
        lowColor: colorLow ?? undefined,
        markerColor: colorMarkers ?? undefined,
        markers: hasEnabledSpreadsheetFlag(markersNode),
        negative: hasEnabledSpreadsheetFlag(negativeNode),
        negativeColor: colorNegative ?? undefined,
        range,
        target,
        type: sparklineType
      });
    }
  }

  return sparklines;
}

function parseSpreadsheetFont(node: Element | null): XlsxResolvedCellStyle["font"] {
  if (!node) {
    return undefined;
  }

  const font: Record<string, unknown> = {};
  const size = getFirstChild(node, "sz")?.getAttribute("val");
  const name = getFirstChild(node, "name")?.getAttribute("val");
  const family = getFirstChild(node, "family")?.getAttribute("val");
  const scheme = getFirstChild(node, "scheme")?.getAttribute("val");
  const charset = getFirstChild(node, "charset")?.getAttribute("val");
  const verticalAlign = getFirstChild(node, "vertAlign")?.getAttribute("val");
  const color = parseSpreadsheetColor(getFirstChild(node, "color"));
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "b"))) {
    font.bold = true;
  }
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "i"))) {
    font.italic = true;
  }
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "strike"))) {
    font.strikethrough = true;
  }
  if (getFirstChild(node, "u")) {
    font.underline = getFirstChild(node, "u")?.getAttribute("val") ?? "single";
  }
  if (size !== null && size !== undefined) {
    font.size = Number(size);
  }
  if (name) {
    font.name = name;
  }
  if (family !== null && family !== undefined) {
    font.family = Number(family);
  }
  if (scheme) {
    font.scheme = scheme;
  }
  if (charset !== null && charset !== undefined) {
    font.charset = Number(charset);
  }
  if (verticalAlign) {
    font.verticalAlign = verticalAlign;
  }
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "shadow"))) {
    font.shadow = true;
  }
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "outline"))) {
    font.outline = true;
  }
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "condense"))) {
    font.condense = true;
  }
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "extend"))) {
    font.extend = true;
  }
  if (color) {
    font.color = color;
  }

  return Object.keys(font).length > 0 ? font : undefined;
}

function parseSpreadsheetFill(node: Element | null): XlsxResolvedCellStyle["fill"] {
  if (!node) {
    return undefined;
  }

  const gradientFill = getFirstChild(node, "gradientFill");
  if (gradientFill) {
    const stops = Array.from(gradientFill.childNodes)
      .filter(isElementNode)
      .filter((child) => child.localName === "stop")
      .map((stopNode) => ({
        color: parseSpreadsheetColor(Array.from(stopNode.childNodes).find(isElementNode) ?? null),
        position: Number(stopNode.getAttribute("position") ?? Number.NaN)
      }))
      .filter((stop) => stop.color && Number.isFinite(stop.position));
    if (stops.length > 0) {
      return {
        degree: Number(gradientFill.getAttribute("degree") ?? 0),
        fillType: "gradient",
        gradientType: gradientFill.getAttribute("type") ?? "linear",
        stops
      };
    }
  }

  const patternFill = getFirstChild(node, "patternFill");
  if (!patternFill) {
    return undefined;
  }

  const patternType = patternFill.getAttribute("patternType") ?? "none";
  const foreground = parseSpreadsheetColor(getFirstChild(patternFill, "fgColor"));
  const background = parseSpreadsheetColor(getFirstChild(patternFill, "bgColor"));
  const solidColor = foreground ?? background;
  if (patternType === "solid" && solidColor) {
    return {
      color: solidColor,
      fillType: "solid"
    };
  }
  // Differential table styles sometimes omit patternType and only set bgColor.
  // Preserve those fills so header/table overrides are not dropped during import.
  if ((patternType === "none" || patternType === "gray125") && (foreground || background)) {
    return {
      background,
      fillType: "pattern",
      foreground,
      patternType
    };
  }
  if (patternType !== "none" && patternType !== "gray125" && (foreground || background)) {
    return {
      background,
      fillType: "pattern",
      foreground,
      patternType
    };
  }

  return undefined;
}

function parseSpreadsheetBorderEdge(node: Element | null) {
  if (!node) {
    return undefined;
  }

  const style = node.getAttribute("style");
  const color = parseSpreadsheetColor(getFirstChild(node, "color"));
  if (!style || style === "none") {
    return undefined;
  }

  return {
    color,
    style
  };
}

function parseSpreadsheetBorder(node: Element | null): XlsxResolvedCellStyle["border"] {
  if (!node) {
    return undefined;
  }

  const border: Record<string, Record<string, unknown>> = {};
  (["top", "right", "bottom", "left", "horizontal", "vertical"] as const).forEach((edge) => {
    const parsedEdge = parseSpreadsheetBorderEdge(getFirstChild(node, edge));
    if (parsedEdge) {
      border[edge] = parsedEdge;
    }
  });

  return Object.keys(border).length > 0 ? border : undefined;
}

function parseSpreadsheetAlignment(node: Element | null): XlsxResolvedCellStyle["alignment"] {
  if (!node) {
    return undefined;
  }

  const alignment: Record<string, unknown> = {};
  const horizontal = node.getAttribute("horizontal");
  const vertical = node.getAttribute("vertical");
  const wrapText = node.getAttribute("wrapText");
  const indent = node.getAttribute("indent");
  const shrinkToFit = node.getAttribute("shrinkToFit");
  const textRotation = node.getAttribute("textRotation");
  if (horizontal) {
    alignment.horizontal = horizontal;
  }
  if (vertical) {
    alignment.vertical = vertical;
  }
  if (wrapText !== null) {
    alignment.wrapText = wrapText === "1";
  }
  if (shrinkToFit !== null) {
    alignment.shrinkToFit = shrinkToFit === "1";
  }
  if (indent !== null) {
    alignment.indent = Number(indent);
  }
  if (textRotation !== null) {
    const parsedRotation = Number(textRotation);
    if (Number.isFinite(parsedRotation)) {
      alignment.textRotation = parsedRotation;
    }
  }

  return Object.keys(alignment).length > 0 ? alignment : undefined;
}

function parseDifferentialStyle(node: Element | null): XlsxResolvedCellStyle {
  if (!node) {
    return {};
  }

  const style: XlsxResolvedCellStyle = {};
  const font = parseSpreadsheetFont(getFirstChild(node, "font"));
  const fill = parseSpreadsheetFill(getFirstChild(node, "fill"));
  const border = parseSpreadsheetBorder(getFirstChild(node, "border"));
  const alignment = parseSpreadsheetAlignment(getFirstChild(node, "alignment"));

  if (font) {
    style.font = font;
  }
  if (fill) {
    style.fill = fill;
  }
  if (border) {
    style.border = border;
  }
  if (alignment) {
    style.alignment = alignment;
  }

  return style;
}

function parseResolvedXfStyle(
  xfNode: Element,
  fonts: Array<XlsxResolvedCellStyle["font"]>,
  fills: Array<XlsxResolvedCellStyle["fill"]>,
  borders: Array<XlsxResolvedCellStyle["border"]>,
  checkboxComplementIndices?: Set<number>
) {
  const style: XlsxResolvedCellStyle = {};
  const fontId = Number(xfNode.getAttribute("fontId") ?? Number.NaN);
  const fillId = Number(xfNode.getAttribute("fillId") ?? Number.NaN);
  const borderId = Number(xfNode.getAttribute("borderId") ?? Number.NaN);
  const alignment = parseSpreadsheetAlignment(getFirstChild(xfNode, "alignment"));

  if (Number.isFinite(fontId) && fonts[fontId]) {
    style.font = fonts[fontId];
  }
  if (Number.isFinite(fillId) && fills[fillId]) {
    style.fill = fills[fillId];
  }
  if (Number.isFinite(borderId) && borders[borderId]) {
    style.border = borders[borderId];
  }
  if (alignment) {
    style.alignment = alignment;
  }
  const xfComplementNode = getFirstDescendant(xfNode, "xfComplement");
  const xfComplementIndex = Number(xfComplementNode?.getAttribute("i") ?? Number.NaN);
  if (Number.isFinite(xfComplementIndex) && checkboxComplementIndices?.has(xfComplementIndex)) {
    style.cellControl = { kind: "checkbox" };
  }

  return style;
}

function parseWorkbookStyles(archive: ArchiveEntries) {
  const xml = readArchiveText(archive, "xl/styles.xml");
  if (!xml) {
    return {
      defaultFont: null,
      namedCellStyleByName: {},
      styleById: {},
      tableStyleByName: {}
    };
  }

  const document = parseXml(xml);
  if (!document) {
    return {
      defaultFont: null,
      namedCellStyleByName: {},
      styleById: {},
      tableStyleByName: {}
    };
  }

  const fontsNode = getFirstDescendant(document, "fonts");
  const fillsNode = getFirstDescendant(document, "fills");
  const bordersNode = getFirstDescendant(document, "borders");
  const cellStyleXfsNode = getFirstDescendant(document, "cellStyleXfs");
  const cellStylesNode = getFirstDescendant(document, "cellStyles");
  const cellXfsNode = getFirstDescendant(document, "cellXfs");
  const dxfsNode = getFirstDescendant(document, "dxfs");
  const tableStylesNode = getFirstDescendant(document, "tableStyles");
  if (!cellXfsNode) {
    return {
      defaultFont: null,
      namedCellStyleByName: {},
      styleById: {},
      tableStyleByName: {}
    };
  }

  const checkboxComplementIndices = readFeaturePropertyBagCheckboxComplements(archive);
  const fonts = getChildElements(fontsNode ?? document.documentElement, "font").map((node) => parseSpreadsheetFont(node));
  const fills = getChildElements(fillsNode ?? document.documentElement, "fill").map((node) => parseSpreadsheetFill(node));
  const borders = getChildElements(bordersNode ?? document.documentElement, "border").map((node) => parseSpreadsheetBorder(node));
  const differentialStyles = getChildElements(dxfsNode ?? document.documentElement, "dxf").map((node) => parseDifferentialStyle(node));
  const cellStyleXfs = getChildElements(cellStyleXfsNode ?? document.documentElement, "xf").map(
    (node) => parseResolvedXfStyle(node, fonts, fills, borders, checkboxComplementIndices)
  );
  const namedCellStyleByName: Record<string, XlsxResolvedCellStyle> = {};
  const styleById: Record<number, XlsxResolvedCellStyle> = {};
  const tableStyleByName: Record<string, XlsxTableStyleDefinition> = {};

  getChildElements(cellXfsNode, "xf").forEach((xfNode, index) => {
    styleById[index] = parseResolvedXfStyle(xfNode, fonts, fills, borders, checkboxComplementIndices);
  });

  getChildElements(cellStylesNode ?? document.documentElement, "cellStyle").forEach((cellStyleNode) => {
    const name = cellStyleNode.getAttribute("name");
    const xfId = Number(cellStyleNode.getAttribute("xfId") ?? Number.NaN);
    if (!name || !Number.isFinite(xfId)) {
      return;
    }

    const resolvedStyle = cellStyleXfs[xfId];
    if (resolvedStyle) {
      namedCellStyleByName[name] = resolvedStyle;
    }
  });

  getChildElements(tableStylesNode ?? document.documentElement, "tableStyle").forEach((tableStyleNode) => {
    const name = tableStyleNode.getAttribute("name");
    if (!name) {
      return;
    }

    const elements: XlsxTableStyleDefinition = {};
    getChildElements(tableStyleNode, "tableStyleElement").forEach((elementNode) => {
      const type = elementNode.getAttribute("type");
      const dxfId = Number(elementNode.getAttribute("dxfId") ?? Number.NaN);
      if (!type || !Number.isFinite(dxfId)) {
        return;
      }

      const differentialStyle = differentialStyles[dxfId];
      if (differentialStyle) {
        elements[type] = differentialStyle;
      }
    });
    tableStyleByName[name] = elements;
  });

  const normalFont = (namedCellStyleByName.Normal?.font ?? styleById[0]?.font ?? fonts[0]) as Record<string, unknown> | undefined;
  const defaultFont = normalFont ? {
    family: typeof normalFont.name === "string" ? normalFont.name : undefined,
    sizePt: typeof normalFont.size === "number" ? normalFont.size : undefined
  } : null;

  return {
    defaultFont,
    namedCellStyleByName,
    styleById,
    tableStyleByName
  };
}

function parseWorkbookTableMetadata(
  archive: ArchiveEntries,
  workbookSheets: WorkbookSheetInfo[]
) {
  return workbookSheets.map((sheet) => {
    const sheetRelationships = parseRelationships(archive, relsPathForDocument(sheet.path), sheet.path);
    const sheetXml = readArchiveText(archive, sheet.path);
    if (!sheetXml) {
      return [];
    }

    const sheetDocument = parseXml(sheetXml);
    if (!sheetDocument) {
      return [];
    }

    return getLocalElements(sheetDocument, "tablePart").flatMap((tablePartNode) => {
      const relationshipId = getRelationshipId(tablePartNode);
      if (!relationshipId) {
        return [];
      }

      const relationship = sheetRelationships.get(relationshipId);
      if (!relationship) {
        return [];
      }

      const tableXml = readArchiveText(archive, relationship.target);
      if (!tableXml) {
        return [];
      }

      const tableDocument = parseXml(tableXml);
      const tableNode = tableDocument?.documentElement;
      if (!tableNode || tableNode.localName !== "table") {
        return [];
      }

      return [{
        displayName: tableNode.getAttribute("displayName") ?? undefined,
        headerRowCellStyle: tableNode.getAttribute("headerRowCellStyle") ?? undefined,
        name: tableNode.getAttribute("name") ?? undefined,
        reference: tableNode.getAttribute("ref") ?? undefined
      } satisfies WorkbookTableMetadata];
    });
  });
}

function parseSqrefRanges(sqref: string | null | undefined): XlsxCellRange[] {
  if (!sqref) {
    return [];
  }

  return sqref
    .trim()
    .split(/\s+/)
    .flatMap((reference) => {
      const range = parseA1RangeReference(reference);
      return range ? [range] : [];
    });
}

function parseConditionalFormatValueObject(node: Element | null): XlsxConditionalFormatValueObject | null {
  if (!node) {
    return null;
  }

  const type = node.getAttribute("type");
  if (!type) {
    return null;
  }

  const rawValue = node.getAttribute("val") ?? getFirstChild(node, "f")?.textContent ?? undefined;
  const numericValue = rawValue !== undefined ? Number(rawValue) : Number.NaN;
  return {
    type,
    value: Number.isFinite(numericValue) ? numericValue : undefined
  };
}

function parseSpreadsheetBooleanAttribute(node: Element | null, name: string) {
  if (!node) {
    return undefined;
  }

  const value = node.getAttribute(name);
  if (value === null) {
    return undefined;
  }

  return value !== "0" && value !== "false";
}

function parseStandardConditionalFormatRule(
  cfRuleNode: Element,
  ranges: XlsxCellRange[]
): (XlsxConditionalFormatRule & { id?: string }) | null {
  const type = cfRuleNode.getAttribute("type");
  const rawPriority = Number(cfRuleNode.getAttribute("priority") ?? Number.NaN);
  const priority = Number.isFinite(rawPriority) ? rawPriority : Number.MAX_SAFE_INTEGER;

  if (type === "colorScale") {
    const colorScaleNode = getFirstChild(cfRuleNode, "colorScale");
    if (!colorScaleNode) {
      return null;
    }

    const cfvos = getChildElements(colorScaleNode, "cfvo")
      .map((node) => parseConditionalFormatValueObject(node))
      .filter((value): value is XlsxConditionalFormatValueObject => Boolean(value));
    const colors = getChildElements(colorScaleNode, "color")
      .map((node) => parseSpreadsheetColor(node))
      .filter((value): value is Record<string, unknown> => Boolean(value));
    if (cfvos.length === 0 || colors.length === 0) {
      return null;
    }

    return {
      cfvos,
      colors,
      kind: "colorScale",
      priority,
      ranges
    };
  }

  if (type === "dataBar") {
    const dataBarNode = getFirstChild(cfRuleNode, "dataBar");
    if (!dataBarNode) {
      return null;
    }

    const cfvos = getChildElements(dataBarNode, "cfvo")
      .map((node) => parseConditionalFormatValueObject(node))
      .filter((value): value is XlsxConditionalFormatValueObject => Boolean(value));
    if (cfvos.length === 0) {
      return null;
    }

    const extId = getFirstDescendant(cfRuleNode, "id")?.textContent?.trim() || undefined;
    return {
      cfvos,
      color: parseSpreadsheetColor(getFirstChild(dataBarNode, "color")),
      kind: "dataBar",
      priority,
      ranges,
      id: extId
    };
  }

  if (type === "iconSet") {
    const iconSetNode = getFirstChild(cfRuleNode, "iconSet");
    if (!iconSetNode) {
      return null;
    }

    const iconSetName = iconSetNode.getAttribute("iconSet");
    const cfvos = getChildElements(iconSetNode, "cfvo")
      .map((node) => parseConditionalFormatValueObject(node))
      .filter((value): value is XlsxConditionalFormatValueObject => Boolean(value));
    if (!iconSetName || cfvos.length === 0) {
      return null;
    }

    return {
      cfvos,
      icons: cfvos.map((_, index) => ({
        iconId: index,
        iconSet: iconSetName
      })),
      kind: "iconSet",
      priority,
      ranges,
      reverse: parseSpreadsheetBooleanAttribute(iconSetNode, "reverse"),
      showValue: parseSpreadsheetBooleanAttribute(iconSetNode, "showValue")
    };
  }

  return null;
}

function parseExtendedConditionalFormatRule(
  cfRuleNode: Element,
  ranges: XlsxCellRange[]
): (XlsxConditionalFormatRule & { id?: string }) | null {
  const type = cfRuleNode.getAttribute("type");
  const ruleId = cfRuleNode.getAttribute("id") ?? undefined;
  const rawPriority = Number(cfRuleNode.getAttribute("priority") ?? Number.NaN);
  const priority = Number.isFinite(rawPriority) ? rawPriority : Number.MAX_SAFE_INTEGER;

  if (type === "dataBar") {
    const dataBarNode = getFirstChild(cfRuleNode, "dataBar");
    if (!dataBarNode) {
      return null;
    }

    const cfvos = getChildElements(dataBarNode, "cfvo")
      .map((node) => parseConditionalFormatValueObject(node))
      .filter((value): value is XlsxConditionalFormatValueObject => Boolean(value));
    if (cfvos.length === 0) {
      return null;
    }

    return {
      axisColor: parseSpreadsheetColor(getFirstChild(dataBarNode, "axisColor")),
      border: parseSpreadsheetBooleanAttribute(dataBarNode, "border"),
      borderColor: parseSpreadsheetColor(getFirstChild(dataBarNode, "borderColor")),
      cfvos,
      color: parseSpreadsheetColor(getFirstChild(dataBarNode, "fillColor")),
      gradient: parseSpreadsheetBooleanAttribute(dataBarNode, "gradient"),
      kind: "dataBar",
      maxLength: Number(dataBarNode.getAttribute("maxLength") ?? Number.NaN),
      minLength: Number(dataBarNode.getAttribute("minLength") ?? Number.NaN),
      negativeBarBorderColorSameAsPositive: parseSpreadsheetBooleanAttribute(dataBarNode, "negativeBarBorderColorSameAsPositive"),
      negativeBorderColor: parseSpreadsheetColor(getFirstChild(dataBarNode, "negativeBorderColor")),
      negativeFillColor: parseSpreadsheetColor(getFirstChild(dataBarNode, "negativeFillColor")),
      priority,
      ranges,
      showValue: parseSpreadsheetBooleanAttribute(dataBarNode, "showValue"),
      id: ruleId
    };
  }

  if (type === "iconSet") {
    const iconSetNode = getFirstChild(cfRuleNode, "iconSet");
    if (!iconSetNode) {
      return null;
    }

    const cfvos = getChildElements(iconSetNode, "cfvo")
      .map((node) => parseConditionalFormatValueObject(node))
      .filter((value): value is XlsxConditionalFormatValueObject => Boolean(value));
    const icons = getChildElements(iconSetNode, "cfIcon")
      .map((iconNode) => {
        const iconSet = iconNode.getAttribute("iconSet");
        const rawIconId = Number(iconNode.getAttribute("iconId") ?? Number.NaN);
        if (!iconSet || !Number.isFinite(rawIconId)) {
          return null;
        }

        return {
          iconId: rawIconId,
          iconSet
        } satisfies XlsxConditionalFormatIcon;
      })
      .filter((icon): icon is XlsxConditionalFormatIcon => Boolean(icon));

    if (cfvos.length === 0 || icons.length === 0) {
      return null;
    }

    return {
      cfvos,
      icons,
      kind: "iconSet",
      priority,
      ranges,
      reverse: parseSpreadsheetBooleanAttribute(iconSetNode, "reverse"),
      showValue: parseSpreadsheetBooleanAttribute(iconSetNode, "showValue"),
      id: ruleId
    };
  }

  return null;
}

function mergeConditionalFormatRule(
  baseRule: XlsxConditionalFormatRule & { id?: string },
  extendedRule: XlsxConditionalFormatRule & { id?: string }
) {
  if (baseRule.kind !== extendedRule.kind) {
    return baseRule;
  }

  if (baseRule.kind === "colorScale" && extendedRule.kind === "colorScale") {
    return {
      ...baseRule,
      ...extendedRule,
      cfvos: extendedRule.cfvos.length > 0 ? extendedRule.cfvos : baseRule.cfvos,
      colors: extendedRule.colors.length > 0 ? extendedRule.colors : baseRule.colors,
      priority: Number.isFinite(extendedRule.priority) ? extendedRule.priority : baseRule.priority,
      ranges: extendedRule.ranges.length > 0 ? extendedRule.ranges : baseRule.ranges
    };
  }

  if (baseRule.kind === "dataBar" && extendedRule.kind === "dataBar") {
    const merged: XlsxConditionalDataBarRule & { id?: string } = {
      ...baseRule,
      ...extendedRule,
      axisColor: extendedRule.axisColor ?? baseRule.axisColor,
      border: extendedRule.border ?? baseRule.border,
      cfvos: extendedRule.cfvos.length > 0 ? extendedRule.cfvos : baseRule.cfvos,
      color: extendedRule.color ?? baseRule.color,
      negativeBarBorderColorSameAsPositive: extendedRule.negativeBarBorderColorSameAsPositive ?? baseRule.negativeBarBorderColorSameAsPositive,
      negativeBorderColor: extendedRule.negativeBorderColor ?? baseRule.negativeBorderColor,
      negativeFillColor: extendedRule.negativeFillColor ?? baseRule.negativeFillColor,
      priority: Number.isFinite(extendedRule.priority) ? extendedRule.priority : baseRule.priority,
      ranges: extendedRule.ranges.length > 0 ? extendedRule.ranges : baseRule.ranges
    };
    return merged;
  }

  if (baseRule.kind === "iconSet" && extendedRule.kind === "iconSet") {
    const merged: XlsxConditionalIconSetRule & { id?: string } = {
      ...baseRule,
      ...extendedRule,
      cfvos: extendedRule.cfvos.length > 0 ? extendedRule.cfvos : baseRule.cfvos,
      icons: extendedRule.icons.length > 0 ? extendedRule.icons : baseRule.icons,
      priority: Number.isFinite(extendedRule.priority) ? extendedRule.priority : baseRule.priority,
      ranges: extendedRule.ranges.length > 0 ? extendedRule.ranges : baseRule.ranges
    };
    return merged;
  }

  return baseRule;
}

function parseConditionalFormatRules(document: Document) {
  const standardRules: Array<XlsxConditionalFormatRule & { id?: string }> = [];
  const extendedRules: Array<XlsxConditionalFormatRule & { id?: string }> = [];

  getLocalElements(document, "conditionalFormatting").forEach((conditionalFormattingNode) => {
    const isExtended = conditionalFormattingNode.namespaceURI !== SPREADSHEET_NS;
    const ranges = isExtended
      ? parseSqrefRanges(getFirstChild(conditionalFormattingNode, "sqref")?.textContent ?? "")
      : parseSqrefRanges(conditionalFormattingNode.getAttribute("sqref"));

    getChildElements(conditionalFormattingNode, "cfRule").forEach((cfRuleNode) => {
      const parsedRule = isExtended
        ? parseExtendedConditionalFormatRule(cfRuleNode, ranges)
        : parseStandardConditionalFormatRule(cfRuleNode, ranges);
      if (parsedRule) {
        if (isExtended) {
          extendedRules.push(parsedRule);
        } else {
          standardRules.push(parsedRule);
        }
      }
    });
  });

  const mergedRules: XlsxConditionalFormatRule[] = [];
  const usedExtendedRuleIds = new Set<string>();
  const extendedRulesById = new Map(
    extendedRules
      .filter((rule) => typeof rule.id === "string" && rule.id.length > 0)
      .map((rule) => [rule.id as string, rule])
  );

  standardRules.forEach((rule) => {
    const matchingExtendedRule = rule.id ? extendedRulesById.get(rule.id) : undefined;
    if (matchingExtendedRule) {
      usedExtendedRuleIds.add(rule.id as string);
      mergedRules.push(mergeConditionalFormatRule(rule, matchingExtendedRule));
      return;
    }

    mergedRules.push(rule);
  });

  extendedRules.forEach((rule) => {
    if (rule.id && usedExtendedRuleIds.has(rule.id)) {
      return;
    }

    mergedRules.push(rule);
  });

  return mergedRules
    .map((rule) => {
      const nextRule = { ...rule } as XlsxConditionalFormatRule & { id?: string };
      delete nextRule.id;
      return nextRule;
    })
    .filter((rule) => rule.ranges.length > 0)
    .sort((left, right) => left.priority - right.priority);
}

function parseSheetState(
  archive: ArchiveEntries,
  path: string,
  options?: ParseWorkbookStructureOptions & {
    defaultFont?: {
      family?: string;
      sizePt?: number;
    } | null;
  }
): WorkbookSheetState | null {
  const xml = readArchiveText(archive, path);
  if (!xml) {
    return null;
  }

  const document = parseXml(xml);
  if (!document) {
    return null;
  }

  const includeCachedFormulaValues = options?.includeCachedFormulaValues ?? true;
  const cachedFormulaValues: Record<string, string> = {};
  const conditionalFormatRules = parseConditionalFormatRules(document);
  const sparklines = parseSheetSparklines(document, options?.themePalette);
  const sheetFormatNode = getLocalElements(document, "sheetFormatPr")[0] ?? null;
  const sheetViewNode = getLocalElements(document, "sheetView")[0] ?? null;
  const rowHeightOverridesPx: Record<number, number> = {};
  const colWidthOverridesPx: Record<number, number> = {};
  const rowStyleIds: Record<number, number> = {};
  const colStyleIds: Record<number, number> = {};
  const hiddenRows = new Set<number>();
  const hiddenCols = new Set<number>();
  let hasHorizontalMerges = false;
  let hasVerticalMerges = false;
  let maxHorizontalMergeEndCol = -1;
  let maxVerticalMergeEndRow = -1;
  const columnWidthCharacterWidthPx = measureColumnCharacterWidthPx(
    options?.defaultFont?.family,
    options?.defaultFont?.sizePt
  );

  const defaultRowHeight = Number(sheetFormatNode?.getAttribute("defaultRowHeight") ?? 15);
  const defaultColWidth = Number(
    sheetFormatNode?.getAttribute("defaultColWidth")
    ?? sheetFormatNode?.getAttribute("baseColWidth")
    ?? 8.43
  );
  const rawZoomScale = Number(
    sheetViewNode?.getAttribute("zoomScale")
    ?? sheetViewNode?.getAttribute("zoomScaleNormal")
    ?? Number.NaN
  );
  const zoomScale = Number.isFinite(rawZoomScale) && rawZoomScale > 0
    ? rawZoomScale
    : 100;

  getLocalElements(document, "row").forEach((rowNode) => {
    const rowIndex = Number(rowNode.getAttribute("r") ?? 0) - 1;
    const height = Number(rowNode.getAttribute("ht") ?? Number.NaN);
    const styleId = Number(rowNode.getAttribute("s") ?? Number.NaN);
    const isHidden = (rowNode.getAttribute("hidden") ?? "0") === "1";
    if (rowIndex >= 0 && Number.isFinite(height)) {
      rowHeightOverridesPx[rowIndex] = Math.max(MIN_ROW_HEIGHT_PX, Math.round(height * 1.33));
    }
    if (rowIndex >= 0 && Number.isFinite(styleId)) {
      rowStyleIds[rowIndex] = styleId;
    }
    if (rowIndex >= 0 && isHidden) {
      hiddenRows.add(rowIndex);
    }

    if (includeCachedFormulaValues) {
      getChildElements(rowNode, "c").forEach((cellNode) => {
        const formulaNode = getFirstChild(cellNode, "f");
        const valueNode = getFirstChild(cellNode, "v");
        const cellRef = cellNode.getAttribute("r");
        if (formulaNode && valueNode && cellRef) {
          cachedFormulaValues[cellRef] = valueNode.textContent ?? "";
        }
      });
    }
  });

  getLocalElements(document, "col").forEach((colNode) => {
    const min = Number(colNode.getAttribute("min") ?? 0) - 1;
    const max = Number(colNode.getAttribute("max") ?? 0) - 1;
    const width = Number(colNode.getAttribute("width") ?? Number.NaN);
    const styleId = Number(colNode.getAttribute("style") ?? Number.NaN);
    const isHidden = (colNode.getAttribute("hidden") ?? "0") === "1";
    if (!Number.isFinite(width)) {
      if (!Number.isFinite(styleId)) {
        return;
      }
    }

    for (let col = min; col <= max; col += 1) {
      if (col >= 0) {
        if (Number.isFinite(width)) {
          const widthPx = sheetColumnWidthToPixels(width, columnWidthCharacterWidthPx);
          colWidthOverridesPx[col] = widthPx;
        }
        if (Number.isFinite(styleId)) {
          colStyleIds[col] = styleId;
        }
        if (isHidden) {
          hiddenCols.add(col);
        }
      }
    }
  });

  getLocalElements(document, "mergeCell").forEach((mergeNode) => {
    const reference = mergeNode.getAttribute("ref");
    const range = reference ? parseA1RangeReference(reference) : null;
    if (!range) {
      return;
    }

    if (range.end.col > range.start.col) {
      hasHorizontalMerges = true;
      maxHorizontalMergeEndCol = Math.max(maxHorizontalMergeEndCol, range.end.col);
    }
    if (range.end.row > range.start.row) {
      hasVerticalMerges = true;
      maxVerticalMergeEndRow = Math.max(maxVerticalMergeEndRow, range.end.row);
    }
  });

  return {
    cachedFormulaValues,
    columnWidthCharacterWidthPx,
    colWidthOverridesPx,
    colStyleIds,
    conditionalFormatRules,
    defaultColWidthPx: sheetColumnWidthToPixels(defaultColWidth, columnWidthCharacterWidthPx),
    defaultRowHeightPx: Math.max(MIN_ROW_HEIGHT_PX, Math.round(defaultRowHeight * 1.33)),
    hasHorizontalMerges,
    hasVerticalMerges,
    maxHorizontalMergeEndCol,
    maxVerticalMergeEndRow,
    hiddenCols: [...hiddenCols].sort((left, right) => left - right),
    hiddenRows: [...hiddenRows].sort((left, right) => left - right),
    rowHeightOverridesPx,
    rowStyleIds,
    showGridLines: (sheetViewNode?.getAttribute("showGridLines") ?? "1") !== "0",
    sparklines,
    zoomScale
  };
}

function clampChannel(value: number) {
  return Math.max(0, Math.min(255, Math.round(value)));
}

function normalizeHexColor(value: string) {
  const hex = value.replace(/^#/, "");
  if (hex.length === 8) {
    return `#${hex.slice(2).toLowerCase()}`;
  }
  if (hex.length === 6) {
    return `#${hex.toLowerCase()}`;
  }
  return "#000000";
}

function hexToRgb(hex: string) {
  const normalized = normalizeHexColor(hex);
  return {
    b: Number.parseInt(normalized.slice(5, 7), 16),
    g: Number.parseInt(normalized.slice(3, 5), 16),
    r: Number.parseInt(normalized.slice(1, 3), 16)
  };
}

function rgbToHex(red: number, green: number, blue: number) {
  return `#${[red, green, blue]
    .map((value) => clampChannel(value).toString(16).padStart(2, "0"))
    .join("")}`;
}

function applyDrawingColorTransforms(
  baseColor: string,
  transformNodes: Element[]
): DrawingColor {
  let { r, g, b } = hexToRgb(baseColor);
  let opacity = 1;
  let lumMod = 1;
  let lumOff = 0;

  transformNodes.forEach((node) => {
    const rawValue = Number(node.getAttribute("val") ?? 0);
    const value = Number.isFinite(rawValue) ? rawValue / 100000 : 0;
    switch (node.localName) {
      case "alpha":
        opacity *= value;
        break;
      case "lumMod":
        lumMod *= value;
        break;
      case "lumOff":
        lumOff += value;
        break;
      case "shade":
        r *= value;
        g *= value;
        b *= value;
        break;
      case "tint":
        r = r + (255 - r) * value;
        g = g + (255 - g) * value;
        b = b + (255 - b) * value;
        break;
      default:
        break;
    }
  });

  if (lumMod !== 1 || lumOff !== 0) {
    r = r * lumMod + 255 * lumOff;
    g = g * lumMod + 255 * lumOff;
    b = b * lumMod + 255 * lumOff;
  }

  return {
    color: rgbToHex(r, g, b),
    opacity: Math.max(0, Math.min(1, opacity))
  };
}

function resolveThemeColorName(theme: ThemeState, name: string | null) {
  if (!name) {
    return null;
  }

  const aliases: Record<string, string> = {
    bg1: "bg1",
    bg2: "bg2",
    tx1: "tx1",
    tx2: "tx2"
  };
  const key = aliases[name] ?? name;
  return theme.colors.get(key) ?? null;
}

function resolveColorValue(colorNode: Element | null, theme: ThemeState): DrawingColor | null {
  if (!colorNode) {
    return null;
  }

  let baseColor: string | null = null;
  if (colorNode.localName === "srgbClr") {
    baseColor = normalizeHexColor(colorNode.getAttribute("val") ?? "");
  } else if (colorNode.localName === "schemeClr") {
    baseColor = resolveThemeColorName(theme, colorNode.getAttribute("val"));
  } else if (colorNode.localName === "scrgbClr") {
    const r = Number(colorNode.getAttribute("r") ?? 0) * 255 / 100000;
    const g = Number(colorNode.getAttribute("g") ?? 0) * 255 / 100000;
    const b = Number(colorNode.getAttribute("b") ?? 0) * 255 / 100000;
    baseColor = rgbToHex(r, g, b);
  } else if (colorNode.localName === "sysClr") {
    baseColor = normalizeHexColor(colorNode.getAttribute("lastClr") ?? "");
  }

  if (!baseColor) {
    return null;
  }

  return applyDrawingColorTransforms(baseColor, Array.from(colorNode.childNodes).filter(isElementNode));
}

function resolveFillColor(fillParent: Element | null, theme: ThemeState): DrawingColor | null {
  if (!fillParent) {
    return null;
  }

  const solidFillNode = getFirstChild(fillParent, "solidFill");
  if (!solidFillNode) {
    return null;
  }

  return resolveColorValue(Array.from(solidFillNode.childNodes).filter(isElementNode)[0] ?? null, theme);
}

function resolveTextTypeface(typeface: string | null, theme: ThemeState) {
  if (!typeface) {
    return undefined;
  }

  if (typeface === "+mn-lt" || typeface === "+mn-ea" || typeface === "+mn-cs") {
    return theme.minorLatinFont ?? undefined;
  }
  if (typeface === "+mj-lt" || typeface === "+mj-ea" || typeface === "+mj-cs") {
    return theme.majorLatinFont ?? undefined;
  }

  return typeface;
}

function isThemeTypeface(typeface: string | null | undefined) {
  return Boolean(typeface && typeface.startsWith("+"));
}

function resolvePreferredTextTypeface(node: Element | null, theme: ThemeState) {
  if (!node) {
    return undefined;
  }

  const latin = getFirstChild(node, "latin")?.getAttribute("typeface") ?? null;
  const eastAsian = getFirstChild(node, "ea")?.getAttribute("typeface") ?? null;
  const complexScript = getFirstChild(node, "cs")?.getAttribute("typeface") ?? null;
  const candidates = [latin, eastAsian, complexScript];
  const explicit = candidates.find((candidate) => candidate && !isThemeTypeface(candidate));
  if (explicit) {
    return explicit;
  }

  return resolveTextTypeface(candidates.find(Boolean) ?? null, theme);
}

type ShapeTextStyle = {
  bold?: boolean;
  color?: string;
  fontFamily?: string;
  fontSizePt?: number;
  italic?: boolean;
  underline?: boolean;
};

function parseShapeTextStyle(node: Element | null, theme: ThemeState, fallbackColor?: DrawingColor | null): ShapeTextStyle {
  if (!node) {
    return {
      color: fallbackColor?.color
    };
  }

  const fillColor = resolveFillColor(node, theme) ?? fallbackColor ?? null;
  const underlineValue = node.getAttribute("u");
  return {
    bold: node.getAttribute("b") === "1" || undefined,
    color: fillColor?.color,
    fontFamily: resolvePreferredTextTypeface(node, theme),
    fontSizePt: node.getAttribute("sz") ? Number(node.getAttribute("sz")) / 100 : undefined,
    italic: node.getAttribute("i") === "1" || undefined,
    underline: underlineValue && underlineValue !== "none" ? true : undefined
  };
}

function mergeShapeTextStyles(...styles: Array<ShapeTextStyle | undefined>): ShapeTextStyle {
  return styles.reduce<ShapeTextStyle>((acc, style) => {
    if (!style) {
      return acc;
    }
    return {
      bold: style.bold ?? acc.bold,
      color: style.color ?? acc.color,
      fontFamily: style.fontFamily ?? acc.fontFamily,
      fontSizePt: style.fontSizePt ?? acc.fontSizePt,
      italic: style.italic ?? acc.italic,
      underline: style.underline ?? acc.underline
    };
  }, {});
}

function parseMarker(node: Element | null) {
  if (!node) {
    return null;
  }

  const col = Number(getFirstChild(node, "col")?.textContent ?? 0);
  const row = Number(getFirstChild(node, "row")?.textContent ?? 0);
  const colOffsetEmu = Number(getFirstChild(node, "colOff")?.textContent ?? 0);
  const rowOffsetEmu = Number(getFirstChild(node, "rowOff")?.textContent ?? 0);

  return {
    col,
    colOffsetEmu,
    row,
    rowOffsetEmu
  };
}

function parseSpreadsheetBooleanValue(value: string | null | undefined) {
  if (value === null || value === undefined) {
    return undefined;
  }

  const normalized = value.trim().toLowerCase();
  if (!normalized) {
    return true;
  }

  return !["0", "false", "none", "off", "unchecked"].includes(normalized);
}

function parseSpreadsheetBooleanNode(node: Element | null) {
  if (!node) {
    return undefined;
  }

  return parseSpreadsheetBooleanValue(node.getAttribute("val") ?? node.textContent);
}

function parseFormControlKind(rawType: string | null | undefined): XlsxFormControl["kind"] {
  const normalized = (rawType ?? "").trim().toLowerCase();
  switch (normalized) {
    case "button":
      return "button";
    case "checkbox":
      return "checkbox";
    case "drop":
      return "dropdown";
    case "editbox":
      return "editbox";
    case "gbox":
      return "group-box";
    case "label":
      return "label";
    case "list":
      return "listbox";
    case "radio":
      return "radio";
    case "scroll":
      return "scrollbar";
    case "spin":
      return "spinner";
    default:
      return "unknown";
  }
}

function parseFormControlShapeId(value: string | null | undefined) {
  const match = (value ?? "").match(/(\d+)(?!.*\d)/);
  if (!match) {
    return null;
  }

  const parsed = Number(match[1]);
  return Number.isFinite(parsed) ? parsed : null;
}

function parseCssDeclarationValue(styleText: string | null | undefined, property: string) {
  if (!styleText) {
    return null;
  }

  const pattern = new RegExp(`${property}\\s*:\\s*([^;]+)`, "i");
  const match = pattern.exec(styleText);
  return match?.[1]?.trim() ?? null;
}

function parseControlTextAlign(styleText: string | null | undefined): XlsxFormControl["textAlign"] {
  const value = parseCssDeclarationValue(styleText, "text-align")?.toLowerCase();
  if (value === "center" || value === "right") {
    return value;
  }
  return value === "left" ? "left" : undefined;
}

function parseVmlFontSizePt(value: string | null | undefined) {
  const parsed = Number(value ?? Number.NaN);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return undefined;
  }

  return parsed > 40 ? parsed / 20 : parsed;
}

function normalizeControlLabel(label: string | null | undefined) {
  if (!label) {
    return undefined;
  }

  const normalized = label
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return normalized.length > 0 ? normalized : undefined;
}

function parseAnchor(anchorNode: Element) {
  if (anchorNode.localName === "anchor") {
    const from = parseMarker(getFirstChild(anchorNode, "from"));
    const to = parseMarker(getFirstChild(anchorNode, "to"));
    return from && to ? { from, kind: "two-cell" as const, to } : null;
  }

  if (anchorNode.localName === "twoCellAnchor") {
    const from = parseMarker(getFirstChild(anchorNode, "from"));
    const to = parseMarker(getFirstChild(anchorNode, "to"));
    return from && to ? { from, kind: "two-cell" as const, to } : null;
  }

  if (anchorNode.localName === "oneCellAnchor") {
    const from = parseMarker(getFirstChild(anchorNode, "from"));
    const extNode = getFirstChild(anchorNode, "ext");
    return from && extNode
      ? {
          from,
          kind: "one-cell" as const,
          sizeEmu: {
            cx: Number(extNode.getAttribute("cx") ?? 0),
            cy: Number(extNode.getAttribute("cy") ?? 0)
          }
        }
      : null;
  }

  const positionNode = getFirstChild(anchorNode, "pos");
  const extNode = getFirstChild(anchorNode, "ext");
  return positionNode && extNode
    ? {
        kind: "absolute" as const,
        positionEmu: {
          x: Number(positionNode.getAttribute("x") ?? 0),
          y: Number(positionNode.getAttribute("y") ?? 0)
        },
        sizeEmu: {
          cx: Number(extNode.getAttribute("cx") ?? 0),
          cy: Number(extNode.getAttribute("cy") ?? 0)
        }
      }
    : null;
}

function anchorToRect(anchor: XlsxImage["anchor"]): DrawingRectEmu {
  if (anchor.kind === "absolute") {
    return {
      cx: anchor.sizeEmu.cx,
      cy: anchor.sizeEmu.cy,
      x: anchor.positionEmu.x,
      y: anchor.positionEmu.y
    };
  }

  if (anchor.kind === "one-cell") {
    return {
      cx: anchor.sizeEmu.cx,
      cy: anchor.sizeEmu.cy,
      x: anchor.from.colOffsetEmu,
      y: anchor.from.rowOffsetEmu
    };
  }

  return {
    cx: Math.max(0, (anchor.to.col - anchor.from.col) * DEFAULT_COL_WIDTH_EMU + anchor.to.colOffsetEmu - anchor.from.colOffsetEmu),
    cy: Math.max(0, (anchor.to.row - anchor.from.row) * DEFAULT_ROW_HEIGHT_EMU + anchor.to.rowOffsetEmu - anchor.from.rowOffsetEmu),
    x: anchor.from.colOffsetEmu,
    y: anchor.from.rowOffsetEmu
  };
}

function resolveSheetColumnWidthPx(sheetState: WorkbookSheetState | null, col: number) {
  if (col < 0) {
    return 0;
  }
  return sheetState?.colWidthOverridesPx[col] ?? sheetState?.defaultColWidthPx ?? emuToPixels(DEFAULT_COL_WIDTH_EMU);
}

function resolveSheetRowHeightPx(sheetState: WorkbookSheetState | null, row: number) {
  if (row < 0) {
    return 0;
  }
  return sheetState?.rowHeightOverridesPx[row] ?? sheetState?.defaultRowHeightPx ?? emuToPixels(DEFAULT_ROW_HEIGHT_EMU);
}

function sumSheetColumnWidthsEmu(sheetState: WorkbookSheetState | null, beforeCol: number) {
  let total = 0;
  for (let col = 0; col < beforeCol; col += 1) {
    total += pixelsToEmu(resolveSheetColumnWidthPx(sheetState, col));
  }
  return total;
}

function sumSheetRowHeightsEmu(sheetState: WorkbookSheetState | null, beforeRow: number) {
  let total = 0;
  for (let row = 0; row < beforeRow; row += 1) {
    total += pixelsToEmu(resolveSheetRowHeightPx(sheetState, row));
  }
  return total;
}

function anchorToAbsoluteRect(anchor: XlsxImage["anchor"], sheetState: WorkbookSheetState | null): DrawingRectEmu {
  if (anchor.kind === "absolute") {
    return anchorToRect(anchor);
  }

  if (anchor.kind === "one-cell") {
    return {
      cx: anchor.sizeEmu.cx,
      cy: anchor.sizeEmu.cy,
      x: sumSheetColumnWidthsEmu(sheetState, anchor.from.col) + anchor.from.colOffsetEmu,
      y: sumSheetRowHeightsEmu(sheetState, anchor.from.row) + anchor.from.rowOffsetEmu
    };
  }

  const left = sumSheetColumnWidthsEmu(sheetState, anchor.from.col) + anchor.from.colOffsetEmu;
  const top = sumSheetRowHeightsEmu(sheetState, anchor.from.row) + anchor.from.rowOffsetEmu;
  const right = sumSheetColumnWidthsEmu(sheetState, anchor.to.col) + anchor.to.colOffsetEmu;
  const bottom = sumSheetRowHeightsEmu(sheetState, anchor.to.row) + anchor.to.rowOffsetEmu;

  return {
    cx: Math.max(0, right - left),
    cy: Math.max(0, bottom - top),
    x: left,
    y: top
  };
}

function rectToAbsoluteAnchor(rect: DrawingRectEmu): XlsxImage["anchor"] {
  return {
    kind: "absolute",
    positionEmu: {
      x: rect.x,
      y: rect.y
    },
    sizeEmu: {
      cx: rect.cx,
      cy: rect.cy
    }
  };
}

function parseTransformRect(xfrmNode: Element | null) {
  if (!xfrmNode) {
    return null;
  }

  const offNode = getFirstChild(xfrmNode, "off");
  const extNode = getFirstChild(xfrmNode, "ext");
  if (!offNode || !extNode) {
    return null;
  }

  return {
    cx: Number(extNode.getAttribute("cx") ?? 0),
    cy: Number(extNode.getAttribute("cy") ?? 0),
    flipH: xfrmNode.getAttribute("flipH") === "1",
    flipV: xfrmNode.getAttribute("flipV") === "1",
    rot: Number(xfrmNode.getAttribute("rot") ?? 0) / 60000,
    x: Number(offNode.getAttribute("x") ?? 0),
    y: Number(offNode.getAttribute("y") ?? 0)
  };
}

function applyGroupTransform(rect: DrawingRectEmu, group: GroupTransform): DrawingRectEmu {
  return {
    cx: rect.cx * group.scaleX,
    cy: rect.cy * group.scaleY,
    x: group.x + (rect.x - group.chX) * group.scaleX,
    y: group.y + (rect.y - group.chY) * group.scaleY
  };
}

function parseGroupTransform(
  groupNode: Element,
  parentGroup: GroupTransform | null,
  fallbackAnchor: XlsxImage["anchor"],
  sheetState: WorkbookSheetState | null
) {
  const xfrmNode = getFirstDescendant(getFirstChild(groupNode, "grpSpPr") ?? groupNode, "xfrm");
  const anchorRect = anchorToAbsoluteRect(fallbackAnchor, sheetState);
  const rect = parseTransformRect(xfrmNode) ?? anchorRect;
  const rootRectMatchesAnchorOrigin = Math.abs(rect.x - anchorRect.x) <= EMU_PER_PIXEL
    && Math.abs(rect.y - anchorRect.y) <= EMU_PER_PIXEL;
  const chOffNode = getFirstChild(xfrmNode ?? groupNode, "chOff");
  const chExtNode = getFirstChild(xfrmNode ?? groupNode, "chExt");
  const rootScaleX = rect.cx !== 0 ? anchorRect.cx / rect.cx : 1;
  const rootScaleY = rect.cy !== 0 ? anchorRect.cy / rect.cy : 1;
  const useRectFrameForRoot = !parentGroup && (
    rootScaleX < 0.85
    || rootScaleX > 1.15
    || rootScaleY < 0.85
    || rootScaleY > 1.15
  );
  const absoluteRect = parentGroup
    ? applyGroupTransform(rect, parentGroup)
    : rootRectMatchesAnchorOrigin
      ? rect
      : anchorRect;
  const childRectX = parentGroup
    ? Number(chOffNode?.getAttribute("x") ?? 0)
    : useRectFrameForRoot
      ? rect.x
      : Number(chOffNode?.getAttribute("x") ?? 0);
  const childRectY = parentGroup
    ? Number(chOffNode?.getAttribute("y") ?? 0)
    : useRectFrameForRoot
      ? rect.y
      : Number(chOffNode?.getAttribute("y") ?? 0);
  const childRectCx = parentGroup
    ? Number(chExtNode?.getAttribute("cx") ?? rect.cx)
    : useRectFrameForRoot
      ? rect.cx
      : Number(chExtNode?.getAttribute("cx") ?? rect.cx);
  const childRectCy = parentGroup
    ? Number(chExtNode?.getAttribute("cy") ?? rect.cy)
    : useRectFrameForRoot
      ? rect.cy
      : Number(chExtNode?.getAttribute("cy") ?? rect.cy);
  return {
    chCx: childRectCx,
    chCy: childRectCy,
    chX: childRectX,
    chY: childRectY,
    cx: absoluteRect.cx,
    cy: absoluteRect.cy,
    scaleX: childRectCx !== 0 ? absoluteRect.cx / childRectCx : (parentGroup?.scaleX ?? 1),
    scaleY: childRectCy !== 0 ? absoluteRect.cy / childRectCy : (parentGroup?.scaleY ?? 1),
    x: absoluteRect.x,
    y: absoluteRect.y
  };
}

function anchorFromNodeOrFallback(
  node: Element,
  fallbackAnchor: XlsxImage["anchor"],
  parentGroup: GroupTransform | null
) {
  const xfrmNode = getFirstDescendant(node, "xfrm");
  const rect = parseTransformRect(xfrmNode);
  if (rect) {
    const scaleX = parentGroup?.scaleX ?? 1;
    const scaleY = parentGroup?.scaleY ?? 1;
    return {
      anchor: parentGroup ? rectToAbsoluteAnchor(applyGroupTransform(rect, parentGroup)) : fallbackAnchor,
      flipH: rect.flipH,
      flipV: rect.flipV,
      rotationDeg: rect.rot,
      scaleX,
      scaleY
    };
  }

  return {
    anchor: fallbackAnchor,
    flipH: false,
    flipV: false,
    rotationDeg: 0,
    scaleX: parentGroup?.scaleX ?? 1,
    scaleY: parentGroup?.scaleY ?? 1
  };
}

function mapParagraphAlign(value: string | null | undefined): XlsxShape["paragraphs"][number]["align"] | undefined {
  switch (value) {
    case "ctr":
      return "center";
    case "just":
      return "justify";
    case "r":
      return "right";
    case "l":
      return "left";
    default:
      return undefined;
  }
}

function mapVerticalAnchor(value: string | null | undefined): "bottom" | "middle" | "top" {
  switch (value) {
    case "b":
      return "bottom";
    case "ctr":
      return "middle";
    default:
      return "top";
  }
}

function createImageAnchorNodes(document: XMLDocument) {
  return document.documentElement
    ? Array.from(document.documentElement.childNodes).filter(
        (node): node is Element =>
          isElementNode(node) &&
          (node.localName === "twoCellAnchor" || node.localName === "oneCellAnchor" || node.localName === "absoluteAnchor")
      )
    : [];
}

function createImageSource(bytes: Uint8Array, mimeType: string, objectUrls: string[]) {
  if (typeof URL !== "undefined" && typeof Blob !== "undefined") {
    const objectUrl = URL.createObjectURL(new Blob([cloneBytes(bytes)], { type: mimeType }));
    objectUrls.push(objectUrl);
    return objectUrl;
  }

  let binary = "";
  for (const value of bytes) {
    binary += String.fromCharCode(value);
  }
  const base64 = typeof btoa === "function" ? btoa(binary) : "";
  return `data:${mimeType};base64,${base64}`;
}

function isStrokeOnlyGeometry(geometry: string) {
  return geometry === "line" || geometry === "arc" || geometry === "leftBrace";
}

function parseShapeStroke(node: Element, styleNode: Element | null, theme: ThemeState): XlsxShape["stroke"] | undefined {
  const lineNode = getFirstDescendant(node, "ln");
  const lineRefNode = styleNode ? getFirstChild(styleNode, "lnRef") : null;
  if (!lineNode && !lineRefNode) {
    return undefined;
  }

  if (lineNode && getFirstChild(lineNode, "noFill")) {
    return { none: true };
  }

  const color = resolveFillColor(lineNode, theme)
    ?? resolveColorValue(Array.from(lineRefNode?.childNodes ?? []).filter(isElementNode)[0] ?? null, theme);
  const widthEmu = Number(lineNode?.getAttribute("w") ?? 0);
  return {
    color: color?.color,
    dash: getFirstChild(lineNode ?? node, "prstDash")?.getAttribute("val") ?? undefined,
    headEndType: getFirstChild(lineNode ?? node, "headEnd")?.getAttribute("type") ?? undefined,
    none: false,
    opacity: color?.opacity,
    tailEndType: getFirstChild(lineNode ?? node, "tailEnd")?.getAttribute("type") ?? undefined,
    widthPx: widthEmu > 0 ? emuToPixels(widthEmu) : undefined
  };
}

function parseShapeFill(
  node: Element,
  styleNode: Element | null,
  theme: ThemeState,
  geometry: string
): XlsxShape["fill"] | undefined {
  const shapePropsNode = getFirstChild(node, "spPr") ?? node;
  const noFillNode = getFirstChild(shapePropsNode, "noFill");
  if (noFillNode) {
    return { none: true };
  }

  if (isStrokeOnlyGeometry(geometry)) {
    return { none: true };
  }

  const fillRefNode = styleNode ? getFirstChild(styleNode, "fillRef") : null;
  const fillColor = resolveFillColor(shapePropsNode, theme)
    ?? resolveColorValue(Array.from(fillRefNode?.childNodes ?? []).filter(isElementNode)[0] ?? null, theme);
  if (!fillColor) {
    return undefined;
  }

  return {
    color: fillColor.color,
    none: false,
    opacity: fillColor.opacity
  };
}

function parsePointNode(pointNode: Element | null) {
  if (!pointNode) {
    return null;
  }

  const x = Number(pointNode.getAttribute("x") ?? Number.NaN);
  const y = Number(pointNode.getAttribute("y") ?? Number.NaN);
  return Number.isFinite(x) && Number.isFinite(y) ? { x, y } : null;
}

function parseCustomGeometryPath(shapeNode: Element): ShapeVectorPath | undefined {
  const pathNode = getFirstDescendant(shapeNode, "path");
  if (!pathNode) {
    return undefined;
  }

  const width = Number(pathNode.getAttribute("w") ?? Number.NaN);
  const height = Number(pathNode.getAttribute("h") ?? Number.NaN);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    return undefined;
  }

  const commands: string[] = [];
  let lastPoint: { x: number; y: number } | null = null;
  for (const child of Array.from(pathNode.childNodes).filter(isElementNode)) {
    if (child.localName === "moveTo") {
      const point = parsePointNode(getFirstChild(child, "pt"));
      if (point) {
        commands.push(`M ${point.x} ${point.y}`);
        lastPoint = point;
      }
      continue;
    }

    if (child.localName === "lnTo") {
      const point = parsePointNode(getFirstChild(child, "pt"));
      if (point && (!lastPoint || point.x !== lastPoint.x || point.y !== lastPoint.y)) {
        commands.push(`L ${point.x} ${point.y}`);
        lastPoint = point;
      }
      continue;
    }

    if (child.localName === "cubicBezTo") {
      const points = getChildElements(child, "pt").map(parsePointNode).filter((point): point is { x: number; y: number } => point !== null);
      if (points.length === 3) {
        commands.push(`C ${points[0].x} ${points[0].y} ${points[1].x} ${points[1].y} ${points[2].x} ${points[2].y}`);
        lastPoint = points[2];
      }
      continue;
    }

    if (child.localName === "close") {
      commands.push("Z");
      lastPoint = null;
    }
  }

  if (commands.length === 0) {
    return undefined;
  }

  return {
    path: commands.join(" "),
    viewBox: {
      height,
      width
    }
  };
}

function parseShapeGeometryAdjustments(shapeNode: Element) {
  const presetGeometry = getFirstDescendant(shapeNode, "prstGeom");
  if (!presetGeometry) {
    return undefined;
  }

  const adjustments: Record<string, number> = {};
  getChildElements(getFirstChild(presetGeometry, "avLst"), "gd").forEach((adjustmentNode) => {
    const name = adjustmentNode.getAttribute("name");
    const formula = adjustmentNode.getAttribute("fmla") ?? "";
    const match = formula.match(/^val\s+(-?\d+(?:\.\d+)?)$/);
    if (!name || !match) {
      return;
    }

    const value = Number.parseFloat(match[1] ?? "");
    if (Number.isFinite(value)) {
      adjustments[name] = value;
    }
  });

  return Object.keys(adjustments).length > 0 ? adjustments : undefined;
}

function parseShapeParagraphs(
  shapeNode: Element,
  styleNode: Element | null,
  theme: ThemeState
): XlsxShape["paragraphs"] {
  const txBodyNode = getFirstChild(shapeNode, "txBody");
  if (!txBodyNode) {
    return [];
  }

  const defaultFontRef = styleNode ? getFirstChild(styleNode, "fontRef") : null;
  const defaultFontColor = resolveColorValue(Array.from(defaultFontRef?.childNodes ?? []).filter(isElementNode)[0] ?? null, theme);
  const listStyleNode = getFirstChild(txBodyNode, "lstStyle");
  const paragraphs: XlsxShape["paragraphs"] = [];

  getChildElements(txBodyNode, "p").forEach((paragraphNode) => {
    const paragraphProps = getFirstChild(paragraphNode, "pPr");
    const paragraphLevel = Number(paragraphProps?.getAttribute("lvl") ?? 0);
    const listLevelProps = getFirstChild(listStyleNode ?? txBodyNode, `lvl${paragraphLevel + 1}pPr`);
    const inheritedStyle = mergeShapeTextStyles(
      parseShapeTextStyle(getFirstChild(listLevelProps ?? txBodyNode, "defRPr"), theme, defaultFontColor),
      parseShapeTextStyle(getFirstChild(paragraphProps ?? paragraphNode, "defRPr"), theme, defaultFontColor)
    );
    const runs: XlsxShape["paragraphs"][number]["runs"] = [];
    let sawRenderableChild = false;

    Array.from(paragraphNode.childNodes).filter(isElementNode).forEach((child) => {
      if (child.localName === "br") {
        sawRenderableChild = true;
        runs.push({ text: "\n" });
        return;
      }
      if (child.localName !== "r") {
        return;
      }

      sawRenderableChild = true;
      const text = getFirstChild(child, "t")?.textContent ?? "";
      const runProps = getFirstChild(child, "rPr");
      const runStyle = mergeShapeTextStyles(
        inheritedStyle,
        parseShapeTextStyle(runProps, theme, defaultFontColor)
      );

      runs.push({
        bold: runStyle.bold,
        color: runStyle.color,
        fontFamily: runStyle.fontFamily,
        fontSizePt: runStyle.fontSizePt,
        italic: runStyle.italic,
        text,
        underline: runStyle.underline
      });
    });

    if (runs.length === 0) {
      if (!sawRenderableChild && !getFirstChild(paragraphNode, "endParaRPr")) {
        return;
      }

      runs.push({
        bold: inheritedStyle.bold,
        color: inheritedStyle.color,
        fontFamily: inheritedStyle.fontFamily,
        fontSizePt: inheritedStyle.fontSizePt,
        italic: inheritedStyle.italic,
        text: " ",
        underline: inheritedStyle.underline
      });
    }

    paragraphs.push({
      align: mapParagraphAlign(paragraphProps?.getAttribute("algn")),
      runs
    });
  });

  return paragraphs;
}

function parseTextBox(shapeNode: Element): XlsxShape["textBox"] | undefined {
  const txBodyNode = getFirstChild(shapeNode, "txBody");
  if (!txBodyNode) {
    return undefined;
  }

  const bodyProps = getFirstChild(txBodyNode, "bodyPr");
  const leftInset = emuToPixels(Number(bodyProps?.getAttribute("lIns") ?? 91440));
  const rightInset = emuToPixels(Number(bodyProps?.getAttribute("rIns") ?? 91440));
  const topInset = emuToPixels(Number(bodyProps?.getAttribute("tIns") ?? 45720));
  const bottomInset = emuToPixels(Number(bodyProps?.getAttribute("bIns") ?? 45720));

  return {
    horizontalAlign: bodyProps?.getAttribute("anchorCtr") === "1" ? "center" : "left",
    insetPx: {
      bottom: bottomInset,
      left: leftInset,
      right: rightInset,
      top: topInset
    },
    verticalAlign: mapVerticalAnchor(bodyProps?.getAttribute("anchor"))
  };
}

function getHyperlinkTarget(
  node: Element | null,
  drawingRelationships: Map<string, RelationshipRecord>
) {
  const hyperlinkNode = node ? getFirstDescendant(node, "hlinkClick") : null;
  const hyperlinkTargetNode = hyperlinkNode ?? node;
  const hyperlinkId = hyperlinkTargetNode ? getRelationshipId(hyperlinkTargetNode) : null;
  return hyperlinkId ? drawingRelationships.get(hyperlinkId)?.target ?? undefined : undefined;
}

function parsePictureNode(
  pictureNode: Element,
  fallbackAnchor: XlsxImage["anchor"],
  drawingRelationships: Map<string, RelationshipRecord>,
  archive: ArchiveEntries,
  contentTypes: ContentTypesState,
  objectUrls: string[],
  workbookSheetIndex: number,
  imageId: string,
  zIndex: number,
  parentGroup: GroupTransform | null
) {
  const blipNode = getFirstDescendant(pictureNode, "blip");
  const svgBlipNode = blipNode ? getFirstDescendant(blipNode, "svgBlip") : null;
  const embedId = (svgBlipNode ? getEmbeddedRelationshipId(svgBlipNode) : null) ?? (blipNode ? getEmbeddedRelationshipId(blipNode) : null);
  if (!embedId) {
    return null;
  }

  const mediaRelationship = drawingRelationships.get(embedId);
  if (!mediaRelationship || mediaRelationship.type !== IMAGE_REL_TYPE) {
    return null;
  }

  const mediaBytes = archive[mediaRelationship.target];
  if (!mediaBytes) {
    return null;
  }

  const nonVisualProps = getFirstDescendant(pictureNode, "cNvPr");
  const transform = anchorFromNodeOrFallback(pictureNode, fallbackAnchor, parentGroup);
  return {
    image: {
      anchor: transform.anchor,
      description: nonVisualProps?.getAttribute("descr") ?? undefined,
      hyperlink: getHyperlinkTarget(nonVisualProps ?? pictureNode, drawingRelationships),
      id: imageId,
      mediaPath: mediaRelationship.target,
      mimeType: resolveContentType(contentTypes, mediaRelationship.target),
      name: nonVisualProps?.getAttribute("name") ?? undefined,
      sheetIndex: workbookSheetIndex,
      src: createImageSource(mediaBytes, resolveContentType(contentTypes, mediaRelationship.target), objectUrls),
      workbookSheetIndex,
      zIndex
    } satisfies XlsxImage,
    mediaPath: mediaRelationship.target
  };
}

function parseShapeNode(
  shapeNode: Element,
  fallbackAnchor: XlsxImage["anchor"],
  drawingRelationships: Map<string, RelationshipRecord>,
  theme: ThemeState,
  workbookSheetIndex: number,
  shapeId: string,
  zIndex: number,
  parentGroup: GroupTransform | null
) {
  const nonVisualProps = getFirstDescendant(shapeNode, "cNvPr");
  const styleNode = getFirstChild(shapeNode, "style");
  const transform = anchorFromNodeOrFallback(shapeNode, fallbackAnchor, parentGroup);
  const geometry = getFirstDescendant(shapeNode, "prstGeom")?.getAttribute("prst")
    ?? (getFirstDescendant(shapeNode, "custGeom") ? "custom" : "rect");
  const customPath = geometry === "custom" ? parseCustomGeometryPath(shapeNode) : undefined;
  const geometryAdjustments = parseShapeGeometryAdjustments(shapeNode);
  return {
    anchor: transform.anchor,
    description: nonVisualProps?.getAttribute("descr") ?? undefined,
    fill: parseShapeFill(shapeNode, styleNode, theme, geometry),
    flipH: transform.flipH,
    flipV: transform.flipV,
    geometry,
    geometryAdjustments,
    hyperlink: getHyperlinkTarget(nonVisualProps ?? shapeNode, drawingRelationships),
    id: shapeId,
    name: nonVisualProps?.getAttribute("name") ?? undefined,
    paragraphs: parseShapeParagraphs(shapeNode, styleNode, theme),
    rotationDeg: transform.rotationDeg,
    scaleX: transform.scaleX,
    scaleY: transform.scaleY,
    sheetIndex: workbookSheetIndex,
    svgPath: customPath?.path,
    svgViewBox: customPath?.viewBox,
    stroke: parseShapeStroke(shapeNode, styleNode, theme),
    textBox: parseTextBox(shapeNode),
    workbookSheetIndex,
    zIndex
  } satisfies XlsxShape;
}

function parseAnchorContents(
  anchorNode: Element,
  fallbackAnchor: XlsxImage["anchor"],
  drawingRelationships: Map<string, RelationshipRecord>,
  archive: ArchiveEntries,
  contentTypes: ContentTypesState,
  objectUrls: string[],
  workbookSheetIndex: number,
  theme: ThemeState,
  ids: { image: number; shape: number; z: number },
  imageOriginsById: Map<string, WorkbookImageOrigin>,
  anchorIndex: number,
  sheetState: WorkbookSheetState | null,
  parentGroup: GroupTransform | null = null
) {
  const images: XlsxImage[] = [];
  const shapes: XlsxShape[] = [];
  const mediaPaths = new Set<string>();

  Array.from(anchorNode.childNodes).filter(isElementNode).forEach((child) => {
    if (child.localName === "pic") {
      const imageId = `sheet-${workbookSheetIndex}-${ids.image}`;
      const parsed = parsePictureNode(
        child,
        fallbackAnchor,
        drawingRelationships,
        archive,
        contentTypes,
        objectUrls,
        workbookSheetIndex,
        imageId,
        ids.z++,
        parentGroup
      );
      ids.image += 1;
      if (parsed) {
        images.push(parsed.image);
        mediaPaths.add(parsed.mediaPath);
        imageOriginsById.set(imageId, {
          anchorIndex,
          workbookSheetIndex
        });
      }
      return;
    }

    if (child.localName === "sp") {
      shapes.push(parseShapeNode(
        child,
        fallbackAnchor,
        drawingRelationships,
        theme,
        workbookSheetIndex,
        `shape-${workbookSheetIndex}-${ids.shape++}`,
        ids.z++,
        parentGroup
      ));
      return;
    }

    if (child.localName === "cxnSp") {
      shapes.push(parseShapeNode(
        child,
        fallbackAnchor,
        drawingRelationships,
        theme,
        workbookSheetIndex,
        `shape-${workbookSheetIndex}-${ids.shape++}`,
        ids.z++,
        parentGroup
      ));
      return;
    }

    if (child.localName !== "grpSp") {
      return;
    }

    const nextGroup = parseGroupTransform(
      child,
      parentGroup,
      fallbackAnchor,
      sheetState
    );
    const groupFallbackAnchor = rectToAbsoluteAnchor({
      cx: nextGroup.cx,
      cy: nextGroup.cy,
      x: nextGroup.x,
      y: nextGroup.y
    });
    const parsedGroup = parseAnchorContents(
      child,
      groupFallbackAnchor,
      drawingRelationships,
      archive,
      contentTypes,
      objectUrls,
      workbookSheetIndex,
      theme,
      ids,
      imageOriginsById,
      anchorIndex,
      sheetState,
      nextGroup
    );
    parsedGroup.images.forEach((image) => images.push(image));
    parsedGroup.shapes.forEach((shape) => shapes.push(shape));
    parsedGroup.mediaPaths.forEach((path) => mediaPaths.add(path));
  });

  return {
    images,
    mediaPaths,
    shapes
  };
}

function parseDrawingObjects(
  archive: ArchiveEntries,
  contentTypes: ContentTypesState,
  drawingPath: string,
  objectUrls: string[],
  workbookSheetIndex: number,
  zIndexBase: number,
  theme: ThemeState,
  sheetState: WorkbookSheetState | null,
  imageOriginsById: Map<string, WorkbookImageOrigin>
) {
  const drawingXml = readArchiveText(archive, drawingPath);
  if (!drawingXml) {
    return {
      images: [] as XlsxImage[],
      mediaPaths: [] as string[],
      shapes: [] as XlsxShape[]
    };
  }

  const drawingDocument = parseXml(drawingXml);
  if (!drawingDocument) {
    return {
      images: [] as XlsxImage[],
      mediaPaths: [] as string[],
      shapes: [] as XlsxShape[]
    };
  }

  const drawingRelationships = parseRelationships(archive, relsPathForDocument(drawingPath), drawingPath);
  const images: XlsxImage[] = [];
  const shapes: XlsxShape[] = [];
  const mediaPaths = new Set<string>();
  const anchorNodes = createImageAnchorNodes(drawingDocument);
  const ids = {
    image: zIndexBase,
    shape: zIndexBase,
    z: zIndexBase
  };

  anchorNodes.forEach((anchorNode, anchorIndex) => {
    const anchor = parseAnchor(anchorNode);
    if (!anchor) {
      return;
    }

    const parsed = parseAnchorContents(
      anchorNode,
      anchor,
      drawingRelationships,
      archive,
      contentTypes,
      objectUrls,
      workbookSheetIndex,
      theme,
      ids,
      imageOriginsById,
      anchorIndex,
      sheetState
    );
    parsed.images.forEach((image) => images.push(image));
    parsed.shapes.forEach((shape) => shapes.push(shape));
    parsed.mediaPaths.forEach((path) => mediaPaths.add(path));
  });

  return {
    images,
    mediaPaths: [...mediaPaths],
    shapes
  };
}

function parseSheetFormControlNodes(
  archive: ArchiveEntries,
  sheetPath: string
) {
  const sheetXml = readArchiveText(archive, sheetPath);
  if (!sheetXml) {
    return [] as ParsedSheetFormControl[];
  }

  const sheetDocument = parseXml(sheetXml);
  if (!sheetDocument) {
    return [] as ParsedSheetFormControl[];
  }

  return getLocalElements(sheetDocument, "control").map((controlNode) => ({
    anchor: parseAnchor(getFirstDescendant(controlNode, "anchor") ?? controlNode),
    controlRelationshipId: getRelationshipId(controlNode),
    name: controlNode.getAttribute("name") ?? undefined,
    shapeId: parseFormControlShapeId(controlNode.getAttribute("shapeId"))
  }));
}

function parseCtrlPropDocument(
  archive: ArchiveEntries,
  ctrlPropPath: string
) {
  const xml = readArchiveText(archive, ctrlPropPath);
  if (!xml) {
    return null;
  }

  const document = parseXml(xml);
  const root = document?.documentElement;
  if (!root) {
    return null;
  }

  return {
    checked: parseSpreadsheetBooleanValue(root.getAttribute("checked")),
    linkedCell: root.getAttribute("fmlaLink") ?? undefined,
    objectType: root.getAttribute("objectType") ?? undefined
  } satisfies ParsedCtrlProp;
}

function parseVmlFormControls(
  archive: ArchiveEntries,
  vmlDrawingPath: string
) {
  const xml = readArchiveText(archive, vmlDrawingPath);
  if (!xml) {
    return new Map<number, ParsedVmlFormControl>();
  }

  const document = parseXml(xml);
  if (!document) {
    return new Map<number, ParsedVmlFormControl>();
  }

  const controls = new Map<number, ParsedVmlFormControl>();
  for (const shapeNode of getLocalElements(document, "shape")) {
    const clientDataNode = getFirstChild(shapeNode, "ClientData");
    if (!clientDataNode) {
      continue;
    }

    const shapeId = parseFormControlShapeId(
      shapeNode.getAttributeNS("urn:schemas-microsoft-com:office:office", "spid")
      ?? shapeNode.getAttribute("o:spid")
      ?? shapeNode.getAttribute("spid")
      ?? shapeNode.getAttribute("id")
    );
    if (shapeId === null) {
      continue;
    }

    const styleText = shapeNode.getAttribute("style");
    const textboxNode = getFirstChild(shapeNode, "textbox");
    const fontNode = textboxNode ? getFirstDescendant(textboxNode, "font") : null;
    const textContainerNode = textboxNode ? getFirstDescendant(textboxNode, "div") : null;
    const label = normalizeControlLabel(textboxNode?.textContent);
    const zIndex = Number(parseCssDeclarationValue(styleText, "z-index") ?? Number.NaN);

    controls.set(shapeId, {
      checked: parseSpreadsheetBooleanNode(getFirstChild(clientDataNode, "Checked")),
      fontFamily: fontNode?.getAttribute("face") ?? undefined,
      fontSizePt: parseVmlFontSizePt(fontNode?.getAttribute("size")),
      hidden: (parseCssDeclarationValue(styleText, "visibility") ?? "").toLowerCase() === "hidden",
      label,
      linkedCell: normalizeControlLabel(getFirstChild(clientDataNode, "FmlaLink")?.textContent),
      objectType: clientDataNode.getAttribute("ObjectType") ?? undefined,
      shapeId,
      textAlign: parseControlTextAlign(textContainerNode?.getAttribute("style")),
      textColor: fontNode?.getAttribute("color") ?? undefined,
      zIndex: Number.isFinite(zIndex) ? zIndex : controls.size + 1
    });
  }

  return controls;
}

function parseSheetFormControls(
  archive: ArchiveEntries,
  sheetPath: string,
  sheetRelationships: Map<string, RelationshipRecord>,
  workbookSheetIndex: number,
  zIndexBase: number
) {
  const controlNodes = parseSheetFormControlNodes(archive, sheetPath);
  if (controlNodes.length === 0) {
    return [] as XlsxFormControl[];
  }

  const legacyDrawingRelationship = [...sheetRelationships.values()].find(
    (relationship) => relationship.type === VML_DRAWING_REL_TYPE
  );
  const vmlControlsByShapeId = legacyDrawingRelationship
    ? parseVmlFormControls(archive, legacyDrawingRelationship.target)
    : new Map<number, ParsedVmlFormControl>();
  const parsedControls: XlsxFormControl[] = [];

  controlNodes.forEach((controlNode, index) => {
    if (!controlNode.anchor) {
      return;
    }

    const ctrlPropRelationship = controlNode.controlRelationshipId
      ? sheetRelationships.get(controlNode.controlRelationshipId) ?? null
      : null;
    const ctrlPropPath = ctrlPropRelationship?.type === CTRL_PROP_REL_TYPE
      ? ctrlPropRelationship.target
      : null;
    const ctrlProp = ctrlPropPath
      ? parseCtrlPropDocument(archive, ctrlPropPath)
      : null;
    const vmlControl = controlNode.shapeId !== null
      ? vmlControlsByShapeId.get(controlNode.shapeId) ?? null
      : null;
    const kind = parseFormControlKind(ctrlProp?.objectType ?? vmlControl?.objectType);

    parsedControls.push({
      anchor: controlNode.anchor,
      checked: ctrlProp?.checked ?? vmlControl?.checked,
      fontFamily: vmlControl?.fontFamily,
      fontSizePt: vmlControl?.fontSizePt,
      hidden: vmlControl?.hidden ?? false,
      id: `form-control-${workbookSheetIndex}-${index}`,
      kind,
      label: vmlControl?.label ?? normalizeControlLabel(controlNode.name),
      linkedCell: ctrlProp?.linkedCell ?? vmlControl?.linkedCell,
      name: controlNode.name,
      sheetIndex: workbookSheetIndex,
      textAlign: vmlControl?.textAlign,
      textColor: vmlControl?.textColor,
      workbookSheetIndex,
      zIndex: zIndexBase + (vmlControl?.zIndex ?? index + 1)
    });
  });

  return parsedControls.sort((left, right) => left.zIndex - right.zIndex);
}

export function revokeWorkbookImageAssets(assets: WorkbookImageAssets | null) {
  if (!assets) {
    return;
  }

  for (const objectUrl of assets.objectUrls) {
    if (objectUrl.startsWith("blob:")) {
      URL.revokeObjectURL(objectUrl);
    }
  }
}

function parseWorkbookStructureAssetsFromArchive(
  archive: ArchiveEntries,
  options?: ParseWorkbookStructureOptions
): WorkbookStructureAssets & {
  contentTypes: ContentTypesState;
  theme: ThemeState;
  workbookSheets: WorkbookSheetInfo[];
} {
  const contentTypes = parseContentTypes(archive);
  const workbookSheets = parseWorkbookSheets(archive);
  const theme = parseWorkbookTheme(archive);
  const themePalette = buildThemePalette(theme);
  const { defaultFont, namedCellStyleByName, styleById, tableStyleByName } = parseWorkbookStyles(archive);
  const tableMetadataByWorkbookSheetIndex = parseWorkbookTableMetadata(archive, workbookSheets);
  return {
    contentTypes,
    namedCellStyleByName,
    sheetStatesByWorkbookSheetIndex: workbookSheets.map((sheet) => parseSheetState(archive, sheet.path, {
      ...options,
      defaultFont,
      themePalette
    })),
    styleById,
    tableMetadataByWorkbookSheetIndex,
    tableStyleByName,
    theme,
    themePalette,
    workbookSheets
  };
}

export function parseWorkbookStructureAssets(
  bytes: Uint8Array,
  options?: ParseWorkbookStructureOptions
): WorkbookStructureAssets {
  const archive = unzipSync(bytes);
  const {
    namedCellStyleByName,
    sheetStatesByWorkbookSheetIndex,
    styleById,
    tableMetadataByWorkbookSheetIndex,
    tableStyleByName,
    themePalette
  } = parseWorkbookStructureAssetsFromArchive(archive, options);

  return {
    namedCellStyleByName,
    sheetStatesByWorkbookSheetIndex,
    styleById,
    tableMetadataByWorkbookSheetIndex,
    tableStyleByName,
    themePalette
  };
}

export function parseWorkbookChartStyleAssets(bytes: Uint8Array): WorkbookChartStyleAssets {
  const archive = unzipSync(bytes);
  const {
    themePalette,
    workbookSheets
  } = parseWorkbookStructureAssetsFromArchive(archive);
  const sheetOrigins: Array<WorkbookImageSheetOrigin | null> = [];

  workbookSheets.forEach((sheet, workbookSheetIndex) => {
    const sheetRelationships = parseRelationships(archive, relsPathForDocument(sheet.path), sheet.path);
    const attachments: XlsxImageAttachment[] = [];

    for (const relationship of sheetRelationships.values()) {
      if (relationship.type !== DRAWING_REL_TYPE) {
        continue;
      }

      const drawingPath = relationship.target;
      const drawingRelsPath = relsPathForDocument(drawingPath);
      attachments.push({
        drawingPath,
        drawingRelsPath: archive[drawingRelsPath] ? drawingRelsPath : null,
        mediaPaths: []
      });
    }

    sheetOrigins[workbookSheetIndex] = attachments.length > 0
      ? {
          attachments,
          workbookSheetIndex
        }
      : null;
  });

  return {
    archive,
    sheetOrigins,
    themePalette
  };
}

export function parseWorkbookImageAssets(bytes: Uint8Array): WorkbookImageAssets {
  const archive = unzipSync(bytes);
  const {
    contentTypes,
    namedCellStyleByName,
    sheetStatesByWorkbookSheetIndex,
    styleById,
    tableMetadataByWorkbookSheetIndex,
    tableStyleByName,
    theme,
    themePalette,
    workbookSheets
  } = parseWorkbookStructureAssetsFromArchive(archive);
  const objectUrls: string[] = [];
  const formControlsByWorkbookSheetIndex: XlsxFormControl[][] = [];
  const imagesByWorkbookSheetIndex: XlsxImage[][] = [];
  const shapesByWorkbookSheetIndex: XlsxShape[][] = [];
  const sheetOrigins: Array<WorkbookImageSheetOrigin | null> = [];
  const imageOriginsById = new Map<string, WorkbookImageOrigin>();

  workbookSheets.forEach((sheet, workbookSheetIndex) => {
    const sheetRelationships = parseRelationships(archive, relsPathForDocument(sheet.path), sheet.path);
    const attachments: XlsxImageAttachment[] = [];
    const imageList: XlsxImage[] = [];
    const shapeList: XlsxShape[] = [];
    let zIndexBase = 1;

    for (const relationship of sheetRelationships.values()) {
      if (relationship.type !== DRAWING_REL_TYPE) {
        continue;
      }

      const drawingPath = relationship.target;
      const drawingRelsPath = relsPathForDocument(drawingPath);
      const drawingImages = parseDrawingObjects(
        archive,
        contentTypes,
        drawingPath,
        objectUrls,
        workbookSheetIndex,
        zIndexBase,
        theme,
        sheetStatesByWorkbookSheetIndex[workbookSheetIndex] ?? null,
        imageOriginsById
      );
      imageList.push(...drawingImages.images);
      shapeList.push(...drawingImages.shapes);
      zIndexBase += drawingImages.images.length + drawingImages.shapes.length + 10;
      attachments.push({
        drawingPath,
        drawingRelsPath: archive[drawingRelsPath] ? drawingRelsPath : null,
        mediaPaths: drawingImages.mediaPaths
      });
    }

    const formControlList = parseSheetFormControls(
      archive,
      sheet.path,
      sheetRelationships,
      workbookSheetIndex,
      zIndexBase
    );

    formControlsByWorkbookSheetIndex[workbookSheetIndex] = formControlList;
    imagesByWorkbookSheetIndex[workbookSheetIndex] = imageList;
    shapesByWorkbookSheetIndex[workbookSheetIndex] = shapeList;
    sheetOrigins[workbookSheetIndex] = attachments.length > 0
      ? {
          attachments,
          workbookSheetIndex
        }
      : null;
  });

  return {
    archive,
    formControlsByWorkbookSheetIndex,
    imageOriginsById,
    imagesByWorkbookSheetIndex,
    namedCellStyleByName,
    objectUrls,
    shapesByWorkbookSheetIndex,
    sheetOrigins,
    sheetStatesByWorkbookSheetIndex,
    styleById,
    tableMetadataByWorkbookSheetIndex,
    tableStyleByName,
    themePalette
  };
}

function updateAnchorNode(anchorNode: Element, anchor: XlsxImage["anchor"]) {
  if (anchor.kind === "two-cell") {
    updateMarkerElement(getFirstChild(anchorNode, "from"), anchor.from);
    updateMarkerElement(getFirstChild(anchorNode, "to"), anchor.to);
    return;
  }

  if (anchor.kind === "one-cell") {
    updateMarkerElement(getFirstChild(anchorNode, "from"), anchor.from);
    const extNode = getFirstChild(anchorNode, "ext");
    if (extNode) {
      extNode.setAttribute("cx", String(Math.max(0, Math.round(anchor.sizeEmu.cx))));
      extNode.setAttribute("cy", String(Math.max(0, Math.round(anchor.sizeEmu.cy))));
    }
    return;
  }

  const positionNode = getFirstChild(anchorNode, "pos");
  if (positionNode) {
    positionNode.setAttribute("x", String(Math.max(0, Math.round(anchor.positionEmu.x))));
    positionNode.setAttribute("y", String(Math.max(0, Math.round(anchor.positionEmu.y))));
  }
  const extNode = getFirstChild(anchorNode, "ext");
  if (extNode) {
    extNode.setAttribute("cx", String(Math.max(0, Math.round(anchor.sizeEmu.cx))));
    extNode.setAttribute("cy", String(Math.max(0, Math.round(anchor.sizeEmu.cy))));
  }
}

export function updateWorkbookImageAnchor(
  assets: WorkbookImageAssets,
  imageId: string,
  anchor: XlsxImage["anchor"]
) {
  const origin = assets.imageOriginsById.get(imageId);
  if (!origin) {
    return false;
  }

  const attachments = assets.sheetOrigins[origin.workbookSheetIndex]?.attachments ?? [];
  for (const attachment of attachments) {
    const drawingXml = readArchiveText(assets.archive, attachment.drawingPath);
    if (!drawingXml) {
      continue;
    }

    const drawingDocument = parseXml(drawingXml);
    if (!drawingDocument) {
      continue;
    }

    const anchorNodes = createImageAnchorNodes(drawingDocument);
    const anchorNode = anchorNodes[origin.anchorIndex];
    if (!anchorNode || !getFirstChild(anchorNode, "pic")) {
      continue;
    }

    updateAnchorNode(anchorNode, anchor);
    assets.archive[attachment.drawingPath] = strToU8(serializeXml(drawingDocument));
    const imageList = assets.imagesByWorkbookSheetIndex[origin.workbookSheetIndex] ?? [];
    const imageIndex = imageList.findIndex((image) => image.id === imageId);
    if (imageIndex >= 0) {
      imageList[imageIndex] = {
        ...imageList[imageIndex],
        anchor
      };
    }
    return true;
  }

  return false;
}

function ensureRelationshipsDocument(archive: ArchiveEntries, relsPath: string) {
  const existing = readArchiveText(archive, relsPath);
  if (existing) {
    const parsed = parseXml(existing);
    if (parsed) {
      return parsed;
    }
  }

  return parseXml(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}"></Relationships>`
  );
}

function ensureContentTypesDocument(archive: ArchiveEntries) {
  const existing = readArchiveText(archive, "[Content_Types].xml");
  if (existing) {
    const parsed = parseXml(existing);
    if (parsed) {
      return parsed;
    }
  }

  return parseXml(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${CONTENT_TYPES_NS}"></Types>`
  );
}

function mergeContentTypeForPath(
  targetDocument: XMLDocument,
  originalDocument: XMLDocument | null,
  partPath: string
) {
  const normalizedPartName = `/${normalizeArchivePath(partPath)}`;
  const targetRoot = targetDocument.documentElement;
  if (!targetRoot) {
    return;
  }

  const existingOverride = getLocalElements(targetDocument, "Override").find(
    (node) => node.getAttribute("PartName") === normalizedPartName
  );
  if (!existingOverride && originalDocument) {
    const sourceOverride = getLocalElements(originalDocument, "Override").find(
      (node) => node.getAttribute("PartName") === normalizedPartName
    );
    if (sourceOverride) {
      targetRoot.appendChild(sourceOverride.cloneNode(true));
      return;
    }
  }

  const extension = normalizedPartName.split(".").pop()?.toLowerCase();
  if (!extension) {
    return;
  }

  const existingDefault = getLocalElements(targetDocument, "Default").find(
    (node) => (node.getAttribute("Extension") ?? "").toLowerCase() === extension
  );
  if (existingDefault) {
    return;
  }

  if (originalDocument) {
    const sourceDefault = getLocalElements(originalDocument, "Default").find(
      (node) => (node.getAttribute("Extension") ?? "").toLowerCase() === extension
    );
    if (sourceDefault) {
      targetRoot.appendChild(sourceDefault.cloneNode(true));
      return;
    }
  }

  if (extension === "xml") {
    const defaultNode = targetDocument.createElementNS(CONTENT_TYPES_NS, "Default");
    defaultNode.setAttribute("Extension", extension);
    defaultNode.setAttribute("ContentType", "application/xml");
    targetRoot.appendChild(defaultNode);
    return;
  }

  if (extension === "rels") {
    const defaultNode = targetDocument.createElementNS(CONTENT_TYPES_NS, "Default");
    defaultNode.setAttribute("Extension", extension);
    defaultNode.setAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
    targetRoot.appendChild(defaultNode);
  }
}

function removeDrawingReferences(sheetDocument: XMLDocument, relDocument: XMLDocument) {
  getLocalElements(sheetDocument, "drawing").forEach((node) => node.remove());
  getLocalElements(relDocument, "Relationship")
    .filter((node) => node.getAttribute("Type") === DRAWING_REL_TYPE)
    .forEach((node) => node.remove());
}

function nextRelationshipId(relDocument: XMLDocument) {
  const existingIds = new Set(
    getLocalElements(relDocument, "Relationship")
      .map((node) => node.getAttribute("Id"))
      .filter((value): value is string => Boolean(value))
  );

  let index = 1;
  while (existingIds.has(`rIdReactXlsxImage${index}`)) {
    index += 1;
  }

  return `rIdReactXlsxImage${index}`;
}

function appendSheetDrawingReference(
  sheetDocument: XMLDocument,
  relationshipId: string
) {
  const worksheet = sheetDocument.documentElement;
  if (!worksheet) {
    return;
  }

  const drawingNode = sheetDocument.createElementNS(SPREADSHEET_NS, "drawing");
  drawingNode.setAttributeNS(REL_NS, "r:id", relationshipId);

  const extLst = getFirstChild(worksheet, "extLst");
  if (extLst) {
    worksheet.insertBefore(drawingNode, extLst);
    return;
  }

  worksheet.appendChild(drawingNode);
}

export function mergeWorkbookImageAssets(
  savedBytes: Uint8Array,
  sourceAssets: WorkbookImageAssets | null,
  sheetOrigins: Array<WorkbookImageSheetOrigin | null>
) {
  if (!sourceAssets || sheetOrigins.every((origin) => !origin?.attachments.length)) {
    return cloneBytes(savedBytes);
  }

  try {
    const archive = unzipSync(savedBytes);
    const workbookSheets = parseWorkbookSheets(archive);
    const originalContentTypesDocument = parseXml(readArchiveText(sourceAssets.archive, "[Content_Types].xml") ?? "");
    const targetContentTypesDocument = ensureContentTypesDocument(archive);
    if (!targetContentTypesDocument) {
      return cloneBytes(savedBytes);
    }

    sheetOrigins.forEach((origin, workbookSheetIndex) => {
      if (!origin?.attachments.length) {
        return;
      }

      const currentSheet = workbookSheets[workbookSheetIndex];
      if (!currentSheet) {
        return;
      }

      const sheetXml = readArchiveText(archive, currentSheet.path);
      if (!sheetXml) {
        return;
      }

      const sheetDocument = parseXml(sheetXml);
      const relsPath = relsPathForDocument(currentSheet.path);
      const relDocument = ensureRelationshipsDocument(archive, relsPath);
      if (!sheetDocument || !relDocument) {
        return;
      }

      removeDrawingReferences(sheetDocument, relDocument);

      origin.attachments.forEach((attachment) => {
        const drawingBytes = sourceAssets.archive[attachment.drawingPath];
        if (!drawingBytes) {
          return;
        }

        archive[attachment.drawingPath] = cloneBytes(drawingBytes);
        mergeContentTypeForPath(targetContentTypesDocument, originalContentTypesDocument, attachment.drawingPath);

        if (attachment.drawingRelsPath) {
          const drawingRelsBytes = sourceAssets.archive[attachment.drawingRelsPath];
          if (drawingRelsBytes) {
            archive[attachment.drawingRelsPath] = cloneBytes(drawingRelsBytes);
            mergeContentTypeForPath(targetContentTypesDocument, originalContentTypesDocument, attachment.drawingRelsPath);
          }
        }

        attachment.mediaPaths.forEach((mediaPath) => {
          const mediaBytes = sourceAssets.archive[mediaPath];
          if (!mediaBytes) {
            return;
          }

          archive[mediaPath] = cloneBytes(mediaBytes);
          mergeContentTypeForPath(targetContentTypesDocument, originalContentTypesDocument, mediaPath);
        });

        const relationshipId = nextRelationshipId(relDocument);
        const relationshipNode = relDocument.createElementNS(PKG_REL_NS, "Relationship");
        relationshipNode.setAttribute("Id", relationshipId);
        relationshipNode.setAttribute("Type", DRAWING_REL_TYPE);
        relationshipNode.setAttribute("Target", relativeArchivePath(currentSheet.path, attachment.drawingPath));
        relDocument.documentElement?.appendChild(relationshipNode);
        appendSheetDrawingReference(sheetDocument, relationshipId);
      });

      archive[currentSheet.path] = strToU8(serializeXml(sheetDocument));
      archive[relsPath] = strToU8(serializeXml(relDocument));
      mergeContentTypeForPath(targetContentTypesDocument, originalContentTypesDocument, relsPath);
    });

    const hasDrawingOverride = getLocalElements(targetContentTypesDocument, "Override").some(
      (node) => node.getAttribute("ContentType") === DRAWING_CONTENT_TYPE
    );
    if (!hasDrawingOverride) {
      for (const path of Object.keys(archive)) {
        if (path.startsWith("xl/drawings/") && path.endsWith(".xml")) {
          mergeContentTypeForPath(targetContentTypesDocument, originalContentTypesDocument, path);
        }
      }
    }

    archive["[Content_Types].xml"] = strToU8(serializeXml(targetContentTypesDocument));
    return zipSync(archive, { level: 6 });
  } catch {
    return cloneBytes(savedBytes);
  }
}

export function emuToPixels(value: number) {
  return value / EMU_PER_PIXEL;
}

export function pixelsToEmu(value: number) {
  return value * EMU_PER_PIXEL;
}

export function pxToSheetColumnWidth(widthPx: number) {
  return (Math.max(widthPx, MIN_COL_WIDTH_PX) - 5) / 7;
}

export function resolveSheetColumnWidthPixels(width: number, columnWidthCharacterWidthPx?: number) {
  return sheetColumnWidthToPixels(width, columnWidthCharacterWidthPx);
}

export function resolveSheetRowHeightPixels(height: number) {
  return Math.max(Math.round(height * 1.33), MIN_ROW_HEIGHT_PX);
}

export function resolveRenderedSheetAxisPixels(sizePx: number, showGridLines = true) {
  return Math.max(0, sizePx) + (showGridLines ? resolveDeviceGridlineThicknessPx() : 0);
}

export function resolveContentSheetAxisPixels(sizePx: number, showGridLines = true) {
  return Math.max(0, sizePx - (showGridLines ? resolveDeviceGridlineThicknessPx() : 0));
}

function markerFromOffset(offsetPx: number, getSizePx: (index: number) => number) {
  let remaining = Math.max(0, offsetPx);
  let index = 0;
  while (remaining > 0) {
    const size = Math.max(1, getSizePx(index));
    if (remaining < size) {
      break;
    }
    remaining -= size;
    index += 1;
  }

  return {
    index,
    offsetPx: remaining
  };
}

export function rectToImageAnchor(
  rect: XlsxImageRect,
  currentAnchor: XlsxImage["anchor"],
  options: {
    contentOffsetLeft: number;
    contentOffsetTop: number;
    getColumnWidthPx: (col: number) => number;
    getRowHeightPx: (row: number) => number;
  }
): XlsxImage["anchor"] {
  const contentLeft = Math.max(0, rect.left - options.contentOffsetLeft);
  const contentTop = Math.max(0, rect.top - options.contentOffsetTop);
  const contentRight = Math.max(contentLeft + 1, rect.left + rect.width - options.contentOffsetLeft);
  const contentBottom = Math.max(contentTop + 1, rect.top + rect.height - options.contentOffsetTop);

  if (currentAnchor.kind === "absolute") {
    return {
      kind: "absolute",
      positionEmu: {
        x: pixelsToEmu(contentLeft),
        y: pixelsToEmu(contentTop)
      },
      sizeEmu: {
        cx: pixelsToEmu(rect.width),
        cy: pixelsToEmu(rect.height)
      }
    };
  }

  const fromCol = markerFromOffset(contentLeft, options.getColumnWidthPx);
  const fromRow = markerFromOffset(contentTop, options.getRowHeightPx);
  if (currentAnchor.kind === "one-cell") {
    return {
      from: {
        col: fromCol.index,
        colOffsetEmu: pixelsToEmu(fromCol.offsetPx),
        row: fromRow.index,
        rowOffsetEmu: pixelsToEmu(fromRow.offsetPx)
      },
      kind: "one-cell",
      sizeEmu: {
        cx: pixelsToEmu(rect.width),
        cy: pixelsToEmu(rect.height)
      }
    };
  }

  const toCol = markerFromOffset(contentRight, options.getColumnWidthPx);
  const toRow = markerFromOffset(contentBottom, options.getRowHeightPx);
  return {
    from: {
      col: fromCol.index,
      colOffsetEmu: pixelsToEmu(fromCol.offsetPx),
      row: fromRow.index,
      rowOffsetEmu: pixelsToEmu(fromRow.offsetPx)
    },
    kind: "two-cell",
    to: {
      col: toCol.index,
      colOffsetEmu: pixelsToEmu(toCol.offsetPx),
      row: toRow.index,
      rowOffsetEmu: pixelsToEmu(toRow.offsetPx)
    }
  };
}

export function resizeImageRect(
  rect: XlsxImageRect,
  handle: XlsxImageResizeHandlePosition,
  deltaX: number,
  deltaY: number,
  minimumSize = 16
) {
  let left = rect.left;
  let top = rect.top;
  let width = rect.width;
  let height = rect.height;

  if (handle.includes("w")) {
    left += deltaX;
    width -= deltaX;
  }
  if (handle.includes("e")) {
    width += deltaX;
  }
  if (handle.includes("n")) {
    top += deltaY;
    height -= deltaY;
  }
  if (handle.includes("s")) {
    height += deltaY;
  }

  if (width < minimumSize) {
    if (handle.includes("w")) {
      left -= minimumSize - width;
    }
    width = minimumSize;
  }
  if (height < minimumSize) {
    if (handle.includes("n")) {
      top -= minimumSize - height;
    }
    height = minimumSize;
  }

  return { height, left, top, width };
}
