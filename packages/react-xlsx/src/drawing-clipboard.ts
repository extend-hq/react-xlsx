import { strFromU8, strToU8, unzipSync } from "fflate";
import { parseWorkbookXml, serializeWorkbookXml } from "./workbook-xml";
import type { WorkbookImageAssets } from "./images";
import type { WorkbookChartAssets } from "./charts";
import type { XlsxImageAnchor } from "./types";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const relPath = (path: string) => path.replace(/([^/]+)$/, "_rels/$1.rels");
const resolve = (path: string, target: string) => new URL(target, `https://xlsx.invalid/${path}`).pathname.slice(1);
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
  kind: "image" | "chart" | "shape";
  anchor: string;
  drawingPath: string;
  archive: Record<string, Uint8Array>;
  sizeEmu?: { cx: number; cy: number };
};

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

const DRAWING_CLIPBOARD_TYPES = ["web data/shadow-workbook", "data/shadow-workbook"];
const RANGE_CLIPBOARD_TYPE = "application/x-react-xlsx-range+json";

function convertDrawingNamespace(source: Element): Element {
  const document = source.ownerDocument;
  const node = document.createElementNS(XDR, `xdr:${source.localName}`);
  for (const attr of Array.from(source.attributes)) {
    if (!attr.name.startsWith("xmlns")) node.setAttributeNS(attr.namespaceURI, attr.name, attr.value);
  }
  for (const child of Array.from(source.children)) {
    if (child.localName === "txSp") {
      const body = Array.from(child.children).find(node => node.localName === "txBody");
      if (body) node.appendChild(convertDrawingNamespace(body));
      continue;
    }
    const structural = ["sp", "pic", "graphicFrame", "grpSp", "cxnSp", "nvGraphicFramePr", "nvPicPr", "nvSpPr", "nvGrpSpPr", "nvCxnSpPr", "cNvPr", "cNvGraphicFramePr", "cNvPicPr", "cNvSpPr", "cNvGrpSpPr", "cNvCxnSpPr", "spPr", "grpSpPr", "blipFill", "txBody", "style"].includes(child.localName)
      || (child.localName === "xfrm" && source.localName === "graphicFrame");
    node.appendChild(structural ? convertDrawingNamespace(child) : child.cloneNode(true));
  }
  if (node.localName === "graphicFrame") {
    const transform = Array.from(node.children).find(child => child.localName === "xfrm");
    if (transform) node.insertBefore(transform, node.children[1] ?? null);
  }
  return node;
}

export function decodeExcelDrawingClipboard(input: string | Uint8Array): DrawingClipboard | null {
  if (input.length > 32 * 1024 * 1024) return null;
  let value: Record<string, unknown> = {};
  let bytes: Uint8Array;
  if (typeof input !== "string" && input[0] === 0x50 && input[1] === 0x4b) bytes = input;
  else {
    try {
      const json = (typeof input === "string" ? input : strFromU8(input)).replace(/^\uFEFF/, "").replace(/\0+$/, "");
      const parsed: unknown = JSON.parse(json);
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
  if (!archive[drawingPath]) return null;
  archive[relPath(drawingPath)] ??= strToU8(`<Relationships xmlns="${PKG}"/>`);
  const drawing = parseWorkbookXml(strFromU8(archive[drawingPath]));
  let anchor = elements(drawing.documentElement).find(node => /^(twoCell|oneCell|absolute)Anchor$/.test(node.localName));
  if (!anchor) {
    const canvas = elements(drawing.documentElement).find(node => node.localName === "lockedCanvas");
    const objects = canvas && Array.from(canvas.children).filter(node => ["graphicFrame", "pic", "sp", "cxnSp", "grpSp"].includes(node.localName));
    if (objects?.length !== 1) return null;
    const object = objects[0];
    const extent = elements(object).find(node => node.localName === "ext" && node.parentElement?.localName === "xfrm");
    anchor = parseWorkbookXml(`<xdr:oneCellAnchor xmlns:xdr="${XDR}"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${Number(extent?.getAttribute("cx")) || 4572000}" cy="${Number(extent?.getAttribute("cy")) || 2743200}"/><xdr:clientData/></xdr:oneCellAnchor>`).documentElement;
    anchor.insertBefore(anchor.ownerDocument.importNode(convertDrawingNamespace(object), true), Array.from(anchor.children).at(-1)!);
  }
  const nodes = elements(anchor);
  const grouped = nodes.some(node => node.localName === "grpSp");
  if (grouped && nodes.some(node => node.localName === "graphicFrame")) return null;
  const kind = nodes.some(node => node.localName === "chart") ? "chart"
    : nodes.some(node => ["sp", "cxnSp", "grpSp"].includes(node.localName)) ? "shape"
    : nodes.some(node => node.localName === "pic") ? "image" : null;
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
  return { kind, anchor: serializeWorkbookXml(anchor), drawingPath, archive, sizeEmu };
}

export async function imageDrawingClipboard(blob: Blob): Promise<DrawingClipboard> {
  if (blob.size > 32 * 1024 * 1024) throw new Error("The copied image is too large.");
  const bitmap = await createImageBitmap(blob);
  let png = blob;
  const sizeEmu = { cx: bitmap.width * 9525, cy: bitmap.height * 9525 };
  try {
    if (bitmap.width * bitmap.height > 16_000_000) throw new Error("The copied image is too large.");
    if (blob.type !== "image/png") {
      const canvas = document.createElement("canvas");
      canvas.width = bitmap.width; canvas.height = bitmap.height;
      const context = canvas.getContext("2d");
      if (!context) throw new Error("Image pasting is unavailable.");
      context.drawImage(bitmap, 0, 0);
      png = await new Promise<Blob>((resolve, reject) => canvas.toBlob(value => value ? resolve(value) : reject(new Error("Image pasting failed.")), "image/png"));
    }
  } finally { bitmap.close(); }
  const drawingPath = "xl/drawings/clipboard.xml";
  const archive = {
    "[Content_Types].xml": strToU8(`<Types xmlns="${CT}"><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`),
    [relPath(drawingPath)]: strToU8(`<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${R}/image" Target="../media/clipboard.png"/></Relationships>`),
    "xl/media/clipboard.png": new Uint8Array(await png.arrayBuffer())
  };
  const anchor = `<xdr:oneCellAnchor xmlns:xdr="${XDR}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${R}"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${sizeEmu.cx}" cy="${sizeEmu.cy}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Picture"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${sizeEmu.cx}" cy="${sizeEmu.cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;
  return { kind: "image", drawingPath, anchor, archive, sizeEmu };
}

export async function readBrowserClipboard(transfer?: DataTransfer): Promise<{
  drawing: DrawingClipboard | null;
  structured?: string;
  text?: string;
}> {
  let text = transfer?.getData("text/plain") || undefined;
  let structured = transfer?.getData(RANGE_CLIPBOARD_TYPE) || undefined;
  let html = transfer?.getData("text/html") ?? "";
  const payloads: Array<string | Blob> = DRAWING_CLIPBOARD_TYPES.map(type => transfer?.getData(type)).filter((value): value is string => Boolean(value));
  const images = Array.from(transfer?.items ?? []).filter(item => item.kind === "file" && item.type.startsWith("image/"))
    .map(item => item.getAsFile()).filter((file): file is File => file !== null);
  const previews: Blob[] = [...images];
  if (transfer && !payloads.length && (structured || (text && (!images.length || /<table[\s>]/i.test(html))))) {
    return { drawing: null, structured, text };
  }

  if (typeof navigator !== "undefined" && navigator.clipboard?.read) {
    try {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        for (const type of item.types) {
          if (DRAWING_CLIPBOARD_TYPES.includes(type)) payloads.push(await item.getType(type));
          else if (type === `web ${RANGE_CLIPBOARD_TYPE}` || type === RANGE_CLIPBOARD_TYPE) structured ??= await (await item.getType(type)).text();
          else if (type === "text/plain") text ??= await (await item.getType(type)).text();
          else if (type === "text/html" && !html) html = await (await item.getType(type)).text();
          else if (type.startsWith("image/")) previews.push(await item.getType(type));
        }
      }
    } catch (error) {
      if (!transfer) throw error;
    }
  } else if (!transfer && typeof navigator !== "undefined" && navigator.clipboard?.readText) {
    text = await navigator.clipboard.readText();
  }

  for (const payload of payloads) {
    try {
      const drawing = decodeExcelDrawingClipboard(typeof payload === "string" ? payload : new Uint8Array(await payload.arrayBuffer()));
      if (drawing) return { drawing };
    } catch {}
  }
  if (payloads.length || (!text && !structured && !/<table[\s>]/i.test(html))) {
    for (const preview of previews) {
      try { return { drawing: await imageDrawingClipboard(preview) }; }
      catch {}
    }
  }
  return { drawing: null, structured, text };
}
