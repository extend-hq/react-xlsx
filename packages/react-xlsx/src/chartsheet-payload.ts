import { parseWorkbookXml } from "./workbook-xml";
import { strFromU8 } from "fflate";
import type { WorkbookImageSheetOrigin } from "./images";
import type { XlsxChartsheet } from "./types";

const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

type Relationship = { target: string; type: string };

function readXml(archive: Record<string, Uint8Array>, path: string) {
  const bytes = archive[path];
  if (!bytes) return null;
  const document = parseWorkbookXml(strFromU8(bytes));
  return document.getElementsByTagName("parsererror").length > 0 ? null : document;
}

function relationshipPath(partPath: string) {
  const separator = partPath.lastIndexOf("/");
  return `${partPath.slice(0, separator + 1)}_rels/${partPath.slice(separator + 1)}.rels`;
}

function resolveTarget(partPath: string, target: string) {
  if (/^[a-z][a-z0-9+.-]*:/i.test(target) || target.startsWith("//")) return null;
  const segments = target.startsWith("/") ? [] : partPath.split("/").slice(0, -1);
  for (const segment of target.replace(/\\/g, "/").split("/")) {
    if (!segment || segment === ".") continue;
    if (segment === "..") {
      if (segments.length === 0) return null;
      segments.pop();
    } else {
      segments.push(segment);
    }
  }
  return segments.join("/");
}

function readRelationships(archive: Record<string, Uint8Array>, partPath: string) {
  const relationships = new Map<string, Relationship>();
  const document = readXml(archive, relationshipPath(partPath));
  for (const element of Array.from(document?.getElementsByTagNameNS("*", "Relationship") ?? [])) {
    if (element.getAttribute("TargetMode") === "External") continue;
    const id = element.getAttribute("Id");
    const rawTarget = element.getAttribute("Target");
    const type = element.getAttribute("Type");
    const target = rawTarget ? resolveTarget(partPath, rawTarget) : null;
    if (id && type && target && archive[target]) relationships.set(id, { target, type });
  }
  return relationships;
}

export function collectChartsheetDrawingOrigins(
  archive: Record<string, Uint8Array>,
  chartsheets: XlsxChartsheet[]
) {
  const origins = new Map<number, WorkbookImageSheetOrigin>();
  if (chartsheets.length === 0) return origins;
  const workbookPath = "xl/workbook.xml";
  const workbookDocument = readXml(archive, workbookPath);
  const workbookRelationships = readRelationships(archive, workbookPath);
  for (const sheet of Array.from(workbookDocument?.getElementsByTagNameNS("*", "sheet") ?? [])) {
    const relationshipId = sheet.getAttributeNS(REL_NS, "id") ?? sheet.getAttribute("r:id");
    const relationship = relationshipId ? workbookRelationships.get(relationshipId) : null;
    if (!relationship?.type.endsWith("/chartsheet")) continue;
    const chartsheet = chartsheets.find((entry) => entry.name === sheet.getAttribute("name"));
    if (!chartsheet || chartsheet.workbookSheetIndex == null) continue;
    const sheetDocument = readXml(archive, relationship.target);
    const sheetRelationships = readRelationships(archive, relationship.target);
    const attachments: WorkbookImageSheetOrigin["attachments"] = [];
    for (const drawing of Array.from(sheetDocument?.getElementsByTagNameNS("*", "drawing") ?? [])) {
      const drawingId = drawing.getAttributeNS(REL_NS, "id") ?? drawing.getAttribute("r:id");
      const drawingRelationship = drawingId ? sheetRelationships.get(drawingId) : null;
      if (!drawingRelationship?.type.endsWith("/drawing")) continue;
      const drawingRelsPath = relationshipPath(drawingRelationship.target);
      attachments.push({
        drawingPath: drawingRelationship.target,
        drawingRelsPath: archive[drawingRelsPath] ? drawingRelsPath : null,
        mediaPaths: []
      });
    }
    origins.set(chartsheet.index, { attachments, workbookSheetIndex: chartsheet.workbookSheetIndex });
  }
  return origins;
}
