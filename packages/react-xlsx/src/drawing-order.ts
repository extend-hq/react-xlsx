import { strFromU8, strToU8 } from "fflate";
import { parseWorkbookXml, serializeWorkbookXml } from "./workbook-xml";
import type { WorkbookImageAssets } from "./images";
import type { WorkbookChartAssets } from "./charts";
import type { XlsxDrawingOrder } from "./types";

export function prepareDrawingOrder(
  assets: WorkbookImageAssets, charts: WorkbookChartAssets, kind: "image" | "chart", id: string, order: XlsxDrawingOrder
) {
  const origin = kind === "image" ? assets.imageOriginsById.get(id) : charts.chartOriginsById.get(id);
  if (!origin) return null;
  const attachments = assets.sheetOrigins[origin.workbookSheetIndex]?.attachments;
  if (attachments?.length !== 1) return null;
  const bytes = assets.archive[origin.drawingPath];
  if (!bytes) return null;
  const document = parseWorkbookXml(strFromU8(bytes));
  const anchors = Array.from(document.documentElement.children).filter(node => /^(twoCell|oneCell|absolute)Anchor$/.test(node.localName));
  const chartAnchors = anchors.filter(node => node.getElementsByTagNameNS("*", "chart").length > 0);
  const anchor = kind === "image" ? anchors[origin.anchorIndex] : chartAnchors[origin.anchorIndex];
  const from = anchors.indexOf(anchor);
  if (from < 0 || Array.from(anchor.children).some(node => node.localName === "grpSp")) return null;
  const to = order === "bringToFront" ? anchors.length - 1 : order === "sendToBack" ? 0
    : Math.max(0, Math.min(anchors.length - 1, from + (order === "bringForward" ? 1 : -1)));
  if (from === to) return null;
  const next = [...anchors];
  next.splice(from, 1);
  next.splice(to, 0, anchor);
  return {
    from, to, sheetIndex: origin.workbookSheetIndex,
    commit() {
      const following = anchors.at(-1)?.nextSibling ?? null;
      for (const node of next) document.documentElement.insertBefore(node, following);
      assets.archive[origin.drawingPath] = strToU8(serializeWorkbookXml(document));
      assets.dirtyArchivePaths.add(origin.drawingPath);
    }
  };
}
