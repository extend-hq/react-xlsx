import { strFromU8, strToU8, zipSync, unzipSync } from "fflate";
import { parseWorkbookXml, serializeWorkbookXml } from "./workbook-xml";
import type { WorkbookImageAssets } from "./images";
import type { WorkbookChartAssets } from "./charts";
import type { XlsxImageAnchor } from "./types";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const LC = "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas";
const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const CX = "http://schemas.microsoft.com/office/drawing/2014/chartex";
const relPath = (path: string) => path.replace(/([^/]+)$/, "_rels/$1.rels");
const resolve = (path: string, target: string) => new URL(target, `https://xlsx.invalid/${path}`).pathname.slice(1);
const relativeTarget = (owner: string, target: string) => {
  const from = owner.split("/").slice(0, -1);
  const to = target.split("/");
  while (from.length && from[0] === to[0]) { from.shift(); to.shift(); }
  return [...from.map(() => ".."), ...to].join("/");
};
const elements = (node: Element): Element[] => [node, ...Array.from(node.children).flatMap(elements)];

export function findDrawingAnchor(assets: WorkbookImageAssets, charts: WorkbookChartAssets, kind: "image" | "chart", id: string) {
  const origin = kind === "image" ? assets.imageOriginsById.get(id) : charts.chartOriginsById.get(id);
  if (!origin || !assets.archive[origin.drawingPath]) return null;
  const document = parseWorkbookXml(strFromU8(assets.archive[origin.drawingPath]));
  const anchors = Array.from(document.documentElement.children).filter(node => /^(twoCell|oneCell|absolute)Anchor$/.test(node.localName));
  const candidates = kind === "image" ? anchors : anchors.filter(node => elements(node).some(child => child.localName === "chart"));
  const anchor = candidates[origin.anchorIndex];
  if (!anchor || elements(anchor).some(node => node.localName === "grpSp")) return null;
  return { document, anchor, anchors, origin };
}

export type DrawingClipboard = {
  token: string;
  kind: "image" | "chart";
  anchor: string;
  drawingPath: string;
  archive: Record<string, Uint8Array>;
  sizeEmu?: { cx: number; cy: number };
};

export function captureDrawingClipboard(assets: WorkbookImageAssets, charts: WorkbookChartAssets, kind: "image" | "chart", id: string): DrawingClipboard | null {
  const found = findDrawingAnchor(assets, charts, kind, id);
  if (!found) return null;
  const archive: Record<string, Uint8Array> = {};
  const collect = (path: string) => {
    if (archive[path] || !assets.archive[path]) return;
    archive[path] = assets.archive[path].slice();
    const relationships = assets.archive[relPath(path)];
    if (!relationships) return;
    archive[relPath(path)] = relationships.slice();
    for (const rel of Array.from(parseWorkbookXml(strFromU8(relationships)).documentElement.children)) {
      if (rel.getAttribute("TargetMode") !== "External") collect(resolve(path, rel.getAttribute("Target")!));
    }
  };
  const relationships = assets.archive[relPath(found.origin.drawingPath)];
  if (!relationships) return null;
  const used = new Set(elements(found.anchor).flatMap(node => Array.from(node.attributes).filter(attr => attr.namespaceURI === R).map(attr => attr.value)));
  const rels = parseWorkbookXml(strFromU8(relationships));
  for (const rel of Array.from(rels.documentElement.children)) {
    if (!used.has(rel.getAttribute("Id")!)) rels.documentElement.removeChild(rel);
    else if (rel.getAttribute("TargetMode") !== "External") collect(resolve(found.origin.drawingPath, rel.getAttribute("Target")!));
  }
  archive[relPath(found.origin.drawingPath)] = strToU8(serializeWorkbookXml(rels));
  const themePath = "xl/theme/theme1.xml";
  if (assets.archive[themePath]) {
    archive[themePath] = assets.archive[themePath].slice();
    const themeRel = rels.createElementNS(PKG, "Relationship");
    themeRel.setAttribute("Id", "rIdClipboardTheme"); themeRel.setAttribute("Type", `${R}/theme`); themeRel.setAttribute("Target", `/${themePath}`);
    rels.documentElement.appendChild(themeRel);
    archive[relPath(found.origin.drawingPath)] = strToU8(serializeWorkbookXml(rels));
  }
  archive["[Content_Types].xml"] = assets.archive["[Content_Types].xml"].slice();
  return { token: `react-xlsx-drawing:${crypto.randomUUID()}`, kind, anchor: serializeWorkbookXml(found.anchor), drawingPath: found.origin.drawingPath, archive };
}

export function removeDrawingAnchor(assets: WorkbookImageAssets, charts: WorkbookChartAssets, kind: "image" | "chart", id: string) {
  const found = findDrawingAnchor(assets, charts, kind, id);
  if (!found) return false;
  found.anchor.parentNode!.removeChild(found.anchor);
  assets.archive[found.origin.drawingPath] = strToU8(serializeWorkbookXml(found.document));
  assets.dirtyArchivePaths.add(found.origin.drawingPath);
  return true;
}

export function pasteDrawingClipboard(assets: WorkbookImageAssets, clipboard: DrawingClipboard, sheetIndex: number, name: string, placement?: XlsxImageAnchor) {
  const suffix = crypto.randomUUID().replaceAll("-", "");
  const origin = assets.sheetOrigins[sheetIndex] ??= { workbookSheetIndex: sheetIndex, attachments: [] };
  const attachment = origin.attachments[0] ?? { drawingPath: `xl/drawings/drawing-${suffix}.xml`, drawingRelsPath: null, mediaPaths: [] };
  if (!origin.attachments.length) origin.attachments.push(attachment);
  const path = attachment.drawingPath;
  const document = assets.archive[path] ? parseWorkbookXml(strFromU8(assets.archive[path])) : parseWorkbookXml(`<xdr:wsDr xmlns:xdr="${XDR}" xmlns:r="${R}"/>`);
  const relationships = assets.archive[relPath(path)] ? parseWorkbookXml(strFromU8(assets.archive[relPath(path)])) : parseWorkbookXml(`<Relationships xmlns="${PKG}"/>`);
  const types = parseWorkbookXml(strFromU8(assets.archive["[Content_Types].xml"]));
  const sourceTypes = parseWorkbookXml(strFromU8(clipboard.archive["[Content_Types].xml"]));
  const addType = (part: string, type: string) => {
    const node = types.createElementNS(CT, "Override");
    node.setAttribute("PartName", `/${part}`); node.setAttribute("ContentType", type); types.documentElement.appendChild(node);
  };
  if (!assets.archive[path]) addType(path, "application/vnd.openxmlformats-officedocument.drawing+xml");
  const copies = new Map<string, string>();
  const copyPart = (source: string): string => {
    if (copies.has(source)) return copies.get(source)!;
    const bytes = clipboard.archive[source];
    if (!bytes) throw new Error("A drawing resource is missing.");
    const target = source.replace(/(\.[^/.]+)$/, `-${suffix}$1`);
    copies.set(source, target);
    assets.archive[target] = bytes.slice(); assets.dirtyArchivePaths.add(target);
    const type = Array.from(sourceTypes.documentElement.children).find(node => node.getAttribute("PartName") === `/${source}`)
      ?? Array.from(sourceTypes.documentElement.children).find(node => node.getAttribute("Extension") === source.split(".").at(-1));
    if (type) addType(target, type.getAttribute("ContentType")!);
    if (clipboard.archive[relPath(source)]) {
      const rels = parseWorkbookXml(strFromU8(clipboard.archive[relPath(source)]));
      for (const rel of Array.from(rels.documentElement.children)) if (rel.getAttribute("TargetMode") !== "External") {
        rel.setAttribute("Target", `/${copyPart(resolve(source, rel.getAttribute("Target")!))}`);
      }
      assets.archive[relPath(target)] = strToU8(serializeWorkbookXml(rels)); assets.dirtyArchivePaths.add(relPath(target));
    }
    return target;
  };
  const anchor = document.importNode(parseWorkbookXml(clipboard.anchor).documentElement, true);
  let nextId = Math.max(0, ...elements(document.documentElement).filter(node => node.localName === "cNvPr").map(node => Number(node.getAttribute("id")))) + 1;
  for (const node of elements(anchor)) if (node.localName === "cNvPr") { node.setAttribute("id", String(nextId++)); node.setAttribute("name", name); }
  const usedIds = new Set(Array.from(relationships.documentElement.children).map(node => node.getAttribute("Id")));
  const sourceRels = parseWorkbookXml(strFromU8(clipboard.archive[relPath(clipboard.drawingPath)]));
  const remapped = new Map<string, string>();
  for (const sourceRel of Array.from(sourceRels.documentElement.children)) {
    let index = 1; while (usedIds.has(`rId${index}`)) index++;
    const id = `rId${index}`; usedIds.add(id);
    const rel = relationships.importNode(sourceRel, true); rel.setAttribute("Id", id);
    if (rel.getAttribute("TargetMode") !== "External") rel.setAttribute("Target", `/${copyPart(resolve(clipboard.drawingPath, rel.getAttribute("Target")!))}`);
    relationships.documentElement.appendChild(rel);
    remapped.set(sourceRel.getAttribute("Id")!, id);
  }
  for (const node of elements(anchor)) {
    for (const attr of Array.from(node.attributes)) if (attr.namespaceURI === R && remapped.has(attr.value)) node.setAttributeNS(R, attr.name, remapped.get(attr.value)!);

  }
  if (placement?.kind === "one-cell") {
    const from = placement.from;
    const marker = `<xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>${from.colOffsetEmu}</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>${from.rowOffsetEmu}</xdr:rowOff></xdr:from>`;
    const container = parseWorkbookXml(`<xdr:oneCellAnchor xmlns:xdr="${XDR}">${marker}<xdr:ext cx="${placement.sizeEmu.cx}" cy="${placement.sizeEmu.cy}"/></xdr:oneCellAnchor>`).documentElement;
    for (const node of Array.from(anchor.children)) if (!["from", "to", "pos", "ext"].includes(node.localName)) container.appendChild(node);
    document.documentElement.appendChild(document.importNode(container, true));
  } else document.documentElement.appendChild(anchor);
  for (const [entry, doc] of [[path, document], [relPath(path), relationships], ["[Content_Types].xml", types]] as const) {
    assets.archive[entry] = strToU8(serializeWorkbookXml(doc)); assets.dirtyArchivePaths.add(entry);
  }
}

export const EXCEL_DRAWING_CLIPBOARD_TYPE = "web data/shadow-workbook";
export const INTERNAL_DRAWING_CLIPBOARD_TYPE = "web application/vnd.react-xlsx.drawing+json";

function convertDrawingNamespace(source: Element, namespace: string, prefix: string): Element {
  const document = source.ownerDocument;
  const node = document.createElementNS(namespace, `${prefix}:${source.localName}`);
  for (const attr of Array.from(source.attributes)) {
    if (attr.name.startsWith("xmlns")) continue;
    if (namespace === A && ["macro", "fPublished"].includes(attr.name)) continue;
    node.setAttributeNS(attr.namespaceURI, attr.name, attr.value);
  }
  for (const child of Array.from(source.children)) {
    const structural = ["nvGraphicFramePr", "nvPicPr", "nvSpPr", "cNvPr", "cNvGraphicFramePr", "cNvPicPr", "cNvSpPr", "spPr", "blipFill", "txBody"].includes(child.localName)
      || (child.localName === "xfrm" && source.localName === "graphicFrame");
    node.appendChild(structural ? convertDrawingNamespace(child, namespace, prefix) : child.cloneNode(true));
  }
  if (node.localName === "graphicFrame") {
    const transform = Array.from(node.children).find(child => child.localName === "xfrm");
    if (transform) {
      if (namespace === XDR) node.insertBefore(transform, node.children[1] ?? null);
      else node.appendChild(transform);
    }
  }
  return node;
}

function detachChartData(archive: Record<string, Uint8Array>) {
  for (const [path, bytes] of Object.entries(archive)) {
    if (!/\/charts\/[^/]+\.xml$/.test(path)) continue;
    const document = parseWorkbookXml(strFromU8(bytes));
    if (![C, CX].includes(document.documentElement.namespaceURI ?? "")) continue;
    for (const reference of elements(document.documentElement)) {
      if (reference.namespaceURI === C && ["numRef", "strRef"].includes(reference.localName)) {
        const numeric = reference.localName === "numRef";
        const cache = Array.from(reference.children).find(node => node.localName === (numeric ? "numCache" : "strCache"));
        if (!cache) throw new Error("The chart's source data is unavailable for native copying.");
        let literal: Element;
        if (reference.parentElement?.localName === "tx") {
          const text = elements(cache).find(node => node.localName === "v")?.textContent ?? "";
          if (reference.parentElement.parentElement?.localName === "ser") {
            literal = document.createElementNS(C, "c:v"); literal.textContent = text;
          } else {
            literal = document.createElementNS(C, "c:rich");
            for (const name of ["bodyPr", "lstStyle"]) literal.appendChild(document.createElementNS(A, `a:${name}`));
            const paragraph = document.createElementNS(A, "a:p");
            const run = document.createElementNS(A, "a:r");
            const value = document.createElementNS(A, "a:t"); value.textContent = text;
            run.appendChild(value); paragraph.appendChild(run); literal.appendChild(paragraph);
          }
        } else {
          literal = document.createElementNS(C, numeric ? "c:numLit" : "c:strLit");
          for (const node of Array.from(cache.childNodes)) literal.appendChild(node.cloneNode(true));
        }
        reference.parentNode!.insertBefore(literal, reference);
        reference.parentNode!.removeChild(reference);
      } else if (reference.namespaceURI === C && reference.localName === "multiLvlStrRef") {
        const cache = Array.from(reference.children).find(node => node.localName === "multiLvlStrCache");
        const formula = Array.from(reference.children).find(node => node.localName === "f");
        if (!cache) throw new Error("The chart's category data is unavailable for native copying.");
        if (formula) reference.removeChild(formula);
      } else if (reference.namespaceURI === CX && ["numDim", "strDim"].includes(reference.localName)) {
        if (!Array.from(reference.children).some(node => node.localName === "lvl")) throw new Error("The chart's source data is unavailable for native copying.");
        for (const node of Array.from(reference.children)) if (["f", "nf"].includes(node.localName)) reference.removeChild(node);
      } else if (reference.namespaceURI === CX && reference.localName === "txData") {
        for (const node of Array.from(reference.children)) if (node.localName === "f") reference.removeChild(node);
      }
    }
    for (const node of elements(document.documentElement)) {
      if ([C, CX].includes(node.namespaceURI ?? "") && node.localName === "externalData") node.parentNode!.removeChild(node);
    }
    archive[path] = strToU8(serializeWorkbookXml(document));
  }
}

export function encodeExcelDrawingClipboard(clipboard: DrawingClipboard, format: "binary" | "json" = "json") {
  const archive = { ...clipboard.archive };
  if (format === "binary") detachChartData(archive);
  const anchor = parseWorkbookXml(clipboard.anchor).documentElement;
  const object = elements(anchor).find(node => ["graphicFrame", "pic"].includes(node.localName));
  if (object) {
    const size = clipboard.sizeEmu ?? { cx: 4572000, cy: 2743200 };
    const graphic = convertDrawingNamespace(object, A, "a");
    const transform = elements(graphic).find(node => node.localName === "xfrm");
    const offset = transform && Array.from(transform.children).find(node => node.localName === "off");
    const extent = transform && Array.from(transform.children).find(node => node.localName === "ext");
    offset?.setAttribute("x", "0"); offset?.setAttribute("y", "0");
    extent?.setAttribute("cx", String(size.cx)); extent?.setAttribute("cy", String(size.cy));
    archive[clipboard.drawingPath] = strToU8(`<a:graphic xmlns:a="${A}"><a:graphicData uri="${LC}"><lc:lockedCanvas xmlns:lc="${LC}"><a:nvGrpSpPr><a:cNvPr id="0" name=""/><a:cNvGrpSpPr/></a:nvGrpSpPr><a:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${size.cx}" cy="${size.cy}"/><a:chOff x="0" y="0"/><a:chExt cx="${size.cx}" cy="${size.cy}"/></a:xfrm></a:grpSpPr>${serializeWorkbookXml(graphic)}</lc:lockedCanvas></a:graphicData></a:graphic>`);
  } else {
    archive[clipboard.drawingPath] = strToU8(`<a:graphic xmlns:a="${A}"><a:graphicData uri="${XDR}"><xdr:wsDr xmlns:xdr="${XDR}">${clipboard.anchor}</xdr:wsDr></a:graphicData></a:graphic>`);
  }
  for (const [path, bytes] of Object.entries(archive)) if (/\/theme\/.*\.xml$/.test(path)) {
    const theme = parseWorkbookXml(strFromU8(bytes));
    if (theme.documentElement.localName === "theme") {
      const root = theme.createElementNS(A, "a:clipboardTheme");
      for (const node of Array.from(theme.documentElement.children)) if (["themeElements", "clrMap"].includes(node.localName)) root.appendChild(node.cloneNode(true));
      if (!Array.from(root.children).some(node => node.localName === "clrMap")) {
        const map = theme.createElementNS(A, "a:clrMap");
        for (const [key, value] of Object.entries({ bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2", accent1: "accent1", accent2: "accent2", accent3: "accent3", accent4: "accent4", accent5: "accent5", accent6: "accent6", hlink: "hlink", folHlink: "folHlink" })) map.setAttribute(key, value);
        root.appendChild(map);
      }
      archive[path] = strToU8(serializeWorkbookXml(root));
    }
  }
  archive["_rels/.rels"] = strToU8(`<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${R}/drawing" Target="${clipboard.drawingPath}"/></Relationships>`);
  const types = parseWorkbookXml(strFromU8(archive["[Content_Types].xml"]));
  for (const node of Array.from(types.documentElement.children)) if (node.localName === "Override" && !archive[node.getAttribute("PartName")!.slice(1)]) types.documentElement.removeChild(node);
  archive["[Content_Types].xml"] = strToU8(serializeWorkbookXml(types));
  const paths = new Map(Object.keys(archive).map(path => [path, path.startsWith("xl/") ? `clipboard/${path.slice(3)}` : path]));
  const normalized: Record<string, Uint8Array> = {};
  for (const [path, bytes] of Object.entries(archive)) {
    const target = paths.get(path)!;
    if (path.endsWith(".rels")) {
      const document = parseWorkbookXml(strFromU8(bytes));
      const owner = path === "_rels/.rels" ? "" : path.replace(/\/\_rels\/([^/]+)\.rels$/, "/$1");
      for (const node of Array.from(document.documentElement.children)) if (node.getAttribute("TargetMode") !== "External") {
        const part = resolve(owner, node.getAttribute("Target")!);
        node.setAttribute("Target", relativeTarget(paths.get(owner) ?? owner, paths.get(part) ?? part));
      }
      normalized[target] = strToU8(serializeWorkbookXml(document));
    } else if (path === "[Content_Types].xml") {
      const document = parseWorkbookXml(strFromU8(bytes));
      for (const node of Array.from(document.documentElement.children)) {
        const name = node.getAttribute("PartName")?.slice(1);
        if (name && paths.has(name)) node.setAttribute("PartName", `/${paths.get(name)!}`);
      }
      normalized[target] = strToU8(serializeWorkbookXml(document));
    } else normalized[target] = bytes;
  }
  const bytes = zipSync(normalized);
  if (format === "binary") return new Uint8Array(bytes).buffer;
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 8192) binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
  return JSON.stringify({ shadowWorkbookBase64: btoa(binary), floatingObjectId: `{${crypto.randomUUID().toUpperCase()}}`, floatingObjectType: clipboard.kind === "chart" ? 0 : 1 });
}

export function decodeExcelDrawingClipboard(input: string | Uint8Array): DrawingClipboard | null {
  if (input.length > 32 * 1024 * 1024) return null;
  let value: Record<string, unknown> = {};
  let bytes: Uint8Array;
  if (typeof input !== "string" && input[0] === 0x50 && input[1] === 0x4b) bytes = input;
  else {
    try {
      const parsed: unknown = JSON.parse(typeof input === "string" ? input : strFromU8(input));
      if (!parsed || typeof parsed !== "object" || !("shadowWorkbookBase64" in parsed) || typeof parsed.shadowWorkbookBase64 !== "string") return null;
      value = parsed as Record<string, unknown>;
      bytes = Uint8Array.from(atob(parsed.shadowWorkbookBase64), character => character.charCodeAt(0));
    } catch { return null; }
  }
  let expandedSize = 0;
  const archive = unzipSync(bytes, { filter: entry => {
    expandedSize += entry.originalSize;
    if (expandedSize > 64 * 1024 * 1024) throw new Error("The copied drawing is too large.");
    return true;
  } });
  if (!archive["_rels/.rels"] || !archive["[Content_Types].xml"]) return null;
  const rootRels = parseWorkbookXml(strFromU8(archive["_rels/.rels"]));
  const drawingRel = Array.from(rootRels.documentElement.children).find(node => node.getAttribute("Type") === `${R}/drawing` && node.getAttribute("TargetMode") !== "External");
  if (!drawingRel) return null;
  const drawingPath = resolve("", drawingRel.getAttribute("Target")!);
  if (!archive[drawingPath] || !archive[relPath(drawingPath)]) return null;
  const drawing = parseWorkbookXml(strFromU8(archive[drawingPath]));
  let anchor = elements(drawing.documentElement).find(node => /^(twoCell|oneCell|absolute)Anchor$/.test(node.localName));
  if (!anchor) {
    const canvas = elements(drawing.documentElement).find(node => node.localName === "lockedCanvas");
    const objects = canvas && Array.from(canvas.children).filter(node => ["graphicFrame", "pic"].includes(node.localName));
    if (objects?.length !== 1) return null;
    const object = objects[0];
    const extent = elements(object).find(node => node.localName === "ext" && node.parentElement?.localName === "xfrm");
    anchor = parseWorkbookXml(`<xdr:oneCellAnchor xmlns:xdr="${XDR}"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${Number(extent?.getAttribute("cx")) || 4572000}" cy="${Number(extent?.getAttribute("cy")) || 2743200}"/><xdr:clientData/></xdr:oneCellAnchor>`).documentElement;
    anchor.insertBefore(anchor.ownerDocument.importNode(convertDrawingNamespace(object, XDR, "xdr"), true), Array.from(anchor.children).at(-1)!);
  }
  if (!anchor || elements(anchor).some(node => node.localName === "grpSp")) return null;
  const kind = elements(anchor).some(node => node.localName === "chart") ? "chart" : elements(anchor).some(node => node.localName === "pic") ? "image" : null;
  if (!kind) return null;
  const marker = (name: string) => {
    const node = Array.from(anchor.children).find(child => child.localName === name);
    const number = (key: string) => Number(node && Array.from(node.children).find(child => child.localName === key)?.textContent) || 0;
    return { x: number("col") * 64 * 9525 + number("colOff"), y: number("row") * 20 * 9525 + number("rowOff") };
  };
  const extent = Array.from(anchor.children).find(node => node.localName === "ext");
  const from = marker("from"); const to = marker("to");
  const supplied = "drawingSizeEmu" in value ? value.drawingSizeEmu as { cx?: number; cy?: number } : undefined;
  const sizeEmu = { cx: Math.max(9525, Math.min(100000000, Number(supplied?.cx) || Number(extent?.getAttribute("cx")) || to.x - from.x || 4572000)), cy: Math.max(9525, Math.min(100000000, Number(supplied?.cy) || Number(extent?.getAttribute("cy")) || to.y - from.y || 2743200)) };
  return { token: `react-xlsx-drawing:${crypto.randomUUID()}`, kind, anchor: serializeWorkbookXml(anchor), drawingPath, archive, sizeEmu };
}

export async function drawingPng(source: string | Blob): Promise<Blob> {
  const url = typeof source === "string" ? source : URL.createObjectURL(source);
  try {
    const image = new Image();
    image.src = url;
    await image.decode();
    const canvas = document.createElement("canvas");
    canvas.width = image.naturalWidth; canvas.height = image.naturalHeight;
    const context = canvas.getContext("2d");
    if (!context) throw new Error("Image copying is unavailable.");
    context.drawImage(image, 0, 0);
    return await new Promise<Blob>((resolve, reject) => canvas.toBlob(blob => blob ? resolve(blob) : reject(new Error("Image copying failed.")), "image/png"));
  } finally { if (typeof source !== "string") URL.revokeObjectURL(url); }
}

export async function imageDrawingClipboard(blob: Blob): Promise<DrawingClipboard> {
  const bitmap = await createImageBitmap(blob);
  const sizeEmu = { cx: bitmap.width * 9525, cy: bitmap.height * 9525 };
  bitmap.close();
  const drawingPath = "xl/drawings/clipboard.xml";
  const archive = {
    "[Content_Types].xml": strToU8(`<Types xmlns="${CT}"><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`),
    [relPath(drawingPath)]: strToU8(`<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${R}/image" Target="../media/clipboard.png"/></Relationships>`),
    "xl/media/clipboard.png": new Uint8Array(await blob.arrayBuffer())
  };
  const anchor = `<xdr:oneCellAnchor xmlns:xdr="${XDR}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${R}"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${sizeEmu.cx}" cy="${sizeEmu.cy}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Picture"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${sizeEmu.cx}" cy="${sizeEmu.cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;
  return { token: `react-xlsx-drawing:${crypto.randomUUID()}`, kind: "image", drawingPath, anchor, archive, sizeEmu };
}
