import { strFromU8, strToU8 } from "fflate";
import { parseWorkbookXml, serializeWorkbookXml } from "./workbook-xml";
import type { WorkbookImageAssets } from "./images";
import type { WorkbookChartOrigin } from "./charts";
import type { XlsxChart, XlsxChartReference, XlsxChartSeries } from "./types";
import type { ChartTypeOption } from "./chart-types";

const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const CX = "http://schemas.microsoft.com/office/drawing/2014/chartex";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const classicRelationship = `${R}/chart`;
const modernRelationship = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
const CS = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
const children = (node: Element, name: string) => Array.from(node.children).filter(child => child.localName === name);
const child = (node: Element, name: string) => children(node, name)[0];
const escape = (value: string | number) => String(value).replace(/[&<>"']/g, character => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;" })[character]!);

function dataXml(name: string, reference: XlsxChartReference | null | undefined, values: Array<number | string | null>, numeric: boolean) {
  const type = numeric ? "num" : "str";
  const points = values.map((value, index) => value == null ? "" : `<c:pt idx="${index}"><c:v>${escape(numeric ? Number(value) || 0 : value)}</c:v></c:pt>`).join("");
  const cache = `${numeric ? "<c:formatCode>General</c:formatCode>" : ""}<c:ptCount val="${values.length}"/>${points}`;
  return `<c:${name}>${reference?.formula
    ? `<c:${type}Ref><c:f>${escape(reference.formula.replace(/^=/, ""))}</c:f><c:${type}Cache>${cache}</c:${type}Cache></c:${type}Ref>`
    : `<c:${type}Lit>${cache}</c:${type}Lit>`}</c:${name}>`;
}

function seriesXml(series: XlsxChartSeries, index: number, option: ChartTypeOption, original?: Element) {
  const xy = option.node === "scatterChart" || option.node === "bubbleChart";
  const surface = option.node.startsWith("surface");
  const marker = ["lineChart", "line3DChart", "scatterChart", "radarChart", "stockChart"].includes(option.node);
  const originalTitle = original && child(original, "tx");
  const title = originalTitle ? serializeWorkbookXml(originalTitle) : series.name ? `<c:tx><c:v>${escape(series.name)}</c:v></c:tx>` : "";
  const appearance = original && child(original, "spPr");
  let style = appearance ? serializeWorkbookXml(appearance) : "";
  if (option.scatterStyle === "marker") {
    const node = appearance?.cloneNode(true) as Element | undefined
      ?? parseWorkbookXml(`<c:spPr xmlns:c="${C}" xmlns:a="${A}"/>`).documentElement;
    const line = child(node, "ln");
    if (line) node.removeChild(line);
    const noLine = parseWorkbookXml(`<a:ln xmlns:a="${A}"><a:noFill/></a:ln>`).documentElement;
    node.appendChild(node.ownerDocument.importNode(noLine, true));
    style = serializeWorkbookXml(node);
  }
  const categoryValues = xy && !series.categories.every(value => value == null || typeof value === "number")
    ? series.values.map((_, i) => i + 1) : series.categories;
  const categoryReference = xy && categoryValues !== series.categories ? null : series.categoriesRef;
  const points = !surface && original ? children(original, "dPt").map(node => serializeWorkbookXml(node)).join("") : "";
  const labels = !surface && original && child(original, "dLbls");
  return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${title}${style}`
    + (marker ? `<c:marker><c:symbol val="${option.marker ? "circle" : "none"}"/></c:marker>` : "")
    + points + (labels ? serializeWorkbookXml(labels) : "")
    + dataXml(xy ? "xVal" : "cat", categoryReference, categoryValues, xy || categoryValues.every(value => value == null || typeof value === "number"))
    + dataXml(xy ? "yVal" : "val", series.valuesRef, series.values, true)
    + (option.node === "bubbleChart" ? dataXml("bubbleSize", series.bubbleSizeRef, series.bubbleSizes?.length ? series.bubbleSizes : series.values.map(() => 1), true) + `<c:bubble3D val="${option.bubble3d ? 1 : 0}"/>` : "")
    + (option.node === "scatterChart" || option.node === "lineChart" || option.node === "line3DChart" ? `<c:smooth val="${option.scatterStyle?.startsWith("smooth") ? 1 : 0}"/>` : "")
    + "</c:ser>";
}

function axisXml(kind: "catAx" | "valAx" | "serAx", id: number, crossId: number, position: string, percent = false, hidden = false) {
  return `<c:${kind}><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${hidden ? 1 : 0}"/><c:axPos val="${position}"/>`
    + (kind === "valAx" ? `<c:majorGridlines/><c:numFmt formatCode="${percent ? "0%" : "General"}" sourceLinked="${percent ? 0 : 1}"/>` : "")
    + `<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${crossId}"/><c:crosses val="autoZero"/>`
    + (kind === "catAx" ? '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>' : kind === "valAx" ? '<c:crossBetween val="between"/>' : "") + `</c:${kind}>`;
}

function classicPlotXml(chart: XlsxChart, option: ChartTypeOption, originals: Element[]) {
  const noAxes = /^(pie|ofPie|doughnut)/.test(option.node);
  const xy = option.node === "scatterChart" || option.node === "bubbleChart";
  const thirdAxis = option.node === "line3DChart" || option.node.startsWith("surface");
  const axisIds = noAxes ? [] : thirdAxis ? [100000, 100001, 100002] : [100000, 100001];
  const buildGroup = (selection: XlsxChartSeries[], type: ChartTypeOption, startIndex = 0, groupAxes = axisIds) => {
    let xml = `<c:${type.node}>`;
    if (type.barDir) xml += `<c:barDir val="${type.barDir}"/>`;
    if (type.grouping) xml += `<c:grouping val="${type.grouping}"/>`;
    if (type.scatterStyle) xml += `<c:scatterStyle val="${type.scatterStyle}"/>`;
    if (type.radarStyle) xml += `<c:radarStyle val="${type.radarStyle}"/>`;
    if (type.ofPieType) xml += `<c:ofPieType val="${type.ofPieType}"/>`;
    if (type.node.startsWith("surface")) xml += `<c:wireframe val="${type.wireframe ? 1 : 0}"/>`;
    else if (type.node !== "stockChart") xml += `<c:varyColors val="${noAxes ? 1 : 0}"/>`;
    xml += selection.map((series, index) => seriesXml(series, startIndex + index, type, originals[startIndex + index])).join("");
    if (type.node === "barChart" || type.node === "bar3DChart") xml += '<c:gapWidth val="150"/>';
    if (type.node === "barChart" && type.grouping?.includes("tacked")) xml += '<c:overlap val="100"/>';
    if (type.node === "doughnutChart") xml += '<c:holeSize val="50"/>';
    if (type.node === "bubbleChart") xml += '<c:bubbleScale val="100"/><c:showNegBubbles val="0"/><c:sizeRepresents val="area"/>';
    if (type.node === "stockChart") {
      xml += '<c:hiLowLines/>';
      if (selection.length === 4) xml += '<c:upDownBars><c:gapWidth val="150"/><c:upBars/><c:downBars/></c:upDownBars>';
    }
    xml += groupAxes.map(id => `<c:axId val="${id}"/>`).join("");
    return xml + `</c:${type.node}>`;
  };
  const volume = option.id === "StockVHLC" || option.id === "StockVOHLC";
  let groups = option.combo
    ? buildGroup(chart.series.slice(0, -1), option) + buildGroup(chart.series.slice(-1), { ...option, node: "lineChart", barDir: undefined, grouping: "standard" }, chart.series.length - 1)
    : volume
      ? buildGroup(chart.series.slice(0, 1), { ...option, node: "barChart", grouping: "clustered", barDir: "col" }) + buildGroup(chart.series.slice(1), option, 1, [100003, 100004])
      : buildGroup(chart.series, option);
  if (!noAxes) {
    groups += axisXml(xy ? "valAx" : "catAx", 100000, 100001, option.barDir === "bar" ? "l" : "b")
      + axisXml("valAx", 100001, 100000, option.barDir === "bar" ? "b" : "l", option.grouping === "percentStacked");
    if (thirdAxis) groups += axisXml("serAx", 100002, 100001, "b");
    if (volume) groups += axisXml("catAx", 100003, 100004, "b", false, true) + axisXml("valAx", 100004, 100003, "r");
  }
  return `<c:plotArea><c:layout/>${groups}</c:plotArea>`;
}

function modernChartXml(chart: XlsxChart, option: ChartTypeOption) {
  const dimension = (kind: string, type: string, ref: XlsxChartReference | null | undefined, values: Array<string | number | null>) =>
    `<cx:${kind} type="${type}">${ref?.formula ? `<cx:f>${escape(ref.formula.replace(/^=/, ""))}</cx:f>` : ""}<cx:lvl ptCount="${values.length}"${kind === "numDim" ? ' formatCode="General"' : ""}>${values.map((value, index) => value == null ? "" : `<cx:pt idx="${index}">${escape(value)}</cx:pt>`).join("")}</cx:lvl></cx:${kind}>`;
  const data = chart.series.map((series, index) => `<cx:data id="${index}">`
    + dimension("strDim", "cat", series.categoriesRef, series.categories)
    + dimension("numDim", ["treemap", "sunburst"].includes(option.layout!) ? "size" : "val", series.valuesRef, series.values) + "</cx:data>").join("");
  const histogram = option.id === "Histogram" || option.id === "Pareto";
  const noAxes = ["treemap", "sunburst", "funnel", "regionMap"].includes(option.layout!);
  const series = chart.series.map((entry, index) => `<cx:series uniqueId="{${crypto.randomUUID().toUpperCase()}}" formatIdx="${index}" layoutId="${histogram ? "clusteredColumn" : option.layout}">`
    + (entry.name ? `<cx:tx><cx:txData><cx:v>${escape(entry.name)}</cx:v></cx:txData></cx:tx>` : "")
    + `<cx:dataId val="${index}"/>`
    + (histogram ? '<cx:layoutPr><cx:binning intervalClosed="r"><cx:binCount val="10"/></cx:binning></cx:layoutPr>' : "")
    + (noAxes ? "" : '<cx:axisId val="0"/><cx:axisId val="1"/>') + "</cx:series>"
    + (option.id === "Pareto" ? `<cx:series layoutId="paretoLine" ownerIdx="${index * 2}"><cx:dataId val="${index}"/><cx:axisId val="0"/><cx:axisId val="2"/></cx:series>` : "")).join("");
  const axis = (id: number, category = false) => `<cx:axis id="${id}"><cx:${category ? "catScaling gapWidth=\"1\"" : "valScaling"}/><cx:majorTickMarks type="none"/><cx:minorTickMarks type="none"/><cx:tickLabels/></cx:axis>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cx:chartSpace xmlns:cx="${CX}" xmlns:a="${A}" xmlns:r="${R}"><cx:chartData>${data}</cx:chartData><cx:chart>`
    + (chart.title ? `<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escape(chart.title)}</a:t></a:r></a:p></cx:rich></cx:tx></cx:title>` : "")
    + `<cx:plotArea><cx:plotAreaRegion><cx:plotSurface/>${series}</cx:plotAreaRegion>${noAxes ? "" : axis(0, true) + axis(1) + (option.id === "Pareto" ? axis(2) : "")}</cx:plotArea><cx:legend pos="r" align="ctr" overlay="0"/></cx:chart></cx:chartSpace>`;
}

function ensureModernChartStyles(assets: WorkbookImageAssets, chartPath: string) {
  const path = chartPath.replace(/([^/]+)$/, "_rels/$1.rels");
  const relationships = parseWorkbookXml(assets.archive[path] ? strFromU8(assets.archive[path]) : '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>');
  const types = parseWorkbookXml(strFromU8(assets.archive["[Content_Types].xml"]));
  const names = ["axisTitle", "categoryAxis", "chartArea", "dataLabel", "dataLabelCallout", "dataPoint", "dataPoint3D", "dataPointLine", "dataPointMarker", "dataPointMarkerLayout", "dataPointWireframe", "dataTable", "downBar", "dropLine", "errorBar", "floor", "gridlineMajor", "gridlineMinor", "hiLoLine", "leaderLine", "legend", "plotArea", "plotArea3D", "seriesAxis", "seriesLine", "title", "trendline", "trendlineLabel", "upBar", "valueAxis", "wall"];
  const style = `<cs:chartStyle xmlns:cs="${CS}" xmlns:a="${A}" id="201">${names.map(name => name === "dataPointMarkerLayout" ? '<cs:dataPointMarkerLayout symbol="circle" size="5"/>'
    : `<cs:${name}><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>${name.startsWith("dataPoint") ? '<cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></cs:spPr>' : ""}</cs:${name}>`).join("")}</cs:chartStyle>`;
  const colors = `<cs:colorStyle xmlns:cs="${CS}" xmlns:a="${A}" meth="cycle" id="10">${[1, 2, 3, 4, 5, 6].map(index => `<a:schemeClr val="accent${index}"/>`).join("")}<cs:variation/></cs:colorStyle>`;
  for (const [kind, xml, contentType] of [["chartStyle", style, "chartstyle"], ["chartColorStyle", colors, "chartcolorstyle"]]) {
    const type = `http://schemas.microsoft.com/office/2011/relationships/${kind}`;
    if (Array.from(relationships.documentElement.children).some(node => node.getAttribute("Type") === type)) continue;
    const part = chartPath.replace(/\.xml$/, `-${kind}-${crypto.randomUUID()}.xml`);
    const rel = relationships.createElementNS(relationships.documentElement.namespaceURI, "Relationship");
    rel.setAttribute("Id", `rId${crypto.randomUUID()}`); rel.setAttribute("Type", type); rel.setAttribute("Target", `/${part}`);
    relationships.documentElement.appendChild(rel);
    const override = types.createElementNS(types.documentElement.namespaceURI, "Override");
    override.setAttribute("PartName", `/${part}`); override.setAttribute("ContentType", `application/vnd.ms-office.${contentType}+xml`);
    types.documentElement.appendChild(override);
    assets.archive[part] = strToU8(xml); assets.dirtyArchivePaths.add(part);
  }
  assets.archive[path] = strToU8(serializeWorkbookXml(relationships)); assets.dirtyArchivePaths.add(path);
  assets.archive["[Content_Types].xml"] = strToU8(serializeWorkbookXml(types)); assets.dirtyArchivePaths.add("[Content_Types].xml");
}

export function changeWorkbookChartType(assets: WorkbookImageAssets, origin: WorkbookChartOrigin, chart: XlsxChart, option: ChartTypeOption) {
  if (!origin.chartPath) throw new Error("The chart has no editable source.");
  const old = parseWorkbookXml(strFromU8(assets.archive[origin.chartPath]));
  const isModern = old.documentElement.namespaceURI === CX;
  const nextModern = Boolean(option.layout);
  let xml: string;
  if (nextModern) xml = modernChartXml(chart, option);
  else {
    const document = isModern ? parseWorkbookXml(`<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart>${chart.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escape(chart.title)}</a:t></a:r></a:p></c:rich></c:tx></c:title>` : ""}<c:plotArea/><c:legend><c:legendPos val="r"/></c:legend></c:chart></c:chartSpace>`) : old;
    const chartNode = child(document.documentElement, "chart");
    if (!chartNode) throw new Error("The chart source is invalid.");
    const plot = child(chartNode, "plotArea");
    if (!plot) throw new Error("The chart has no plot area.");
    const originals = Array.from(plot.children).flatMap(node => children(node, "ser"));
    const nextPlot = parseWorkbookXml(`<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}">${classicPlotXml(chart, option, originals)}</c:chartSpace>`).documentElement.firstElementChild!;
    const appearance = child(plot, "spPr");
    if (appearance) nextPlot.appendChild(appearance.cloneNode(true));
    plot.parentNode!.insertBefore(document.importNode(nextPlot, true), plot);
    plot.parentNode!.removeChild(plot);
    for (const node of Array.from(chartNode.children)) if (["view3D", "floor", "sideWall", "backWall"].includes(node.localName)) chartNode.removeChild(node);
    if (option.is3d) {
      const view = parseWorkbookXml(`<c:view3D xmlns:c="${C}"><c:rotX val="15"/><c:rotY val="20"/><c:rAngAx val="1"/></c:view3D>`).documentElement;
      chartNode.insertBefore(document.importNode(view, true), child(chartNode, "plotArea"));
    }
    xml = serializeWorkbookXml(document);
  }
  assets.archive[origin.chartPath] = strToU8(xml);
  assets.dirtyArchivePaths.add(origin.chartPath);
  if (nextModern) ensureModernChartStyles(assets, origin.chartPath);
  if (isModern !== nextModern) {
    const drawing = parseWorkbookXml(strFromU8(assets.archive[origin.drawingPath]));
    const relPath = origin.drawingPath.replace(/([^/]+)$/, "_rels/$1.rels");
    const relationships = parseWorkbookXml(strFromU8(assets.archive[relPath]));
    const rel = Array.from(relationships.documentElement.children).find(node => {
      const target = node.getAttribute("Target");
      return target && new URL(target, `https://xlsx.invalid/${origin.drawingPath}`).pathname.slice(1) === origin.chartPath;
    });
    if (!rel) throw new Error("The chart relationship is missing.");
    rel.setAttribute("Type", nextModern ? modernRelationship : classicRelationship);
    const chartReference = Array.from(drawing.getElementsByTagNameNS(isModern ? CX : C, "chart")).find(node => node.getAttributeNS(R, "id") === rel.getAttribute("Id"));
    if (!chartReference?.parentNode) throw new Error("The chart drawing is missing.");
    const replacement = drawing.createElementNS(nextModern ? CX : C, nextModern ? "cx:chart" : "c:chart");
    replacement.setAttributeNS(R, "r:id", rel.getAttribute("Id")!);
    chartReference.parentElement?.setAttribute("uri", nextModern ? CX : C);
    chartReference.parentNode.insertBefore(replacement, chartReference);
    chartReference.parentNode.removeChild(chartReference);
    const frame = replacement.parentElement?.parentElement?.parentElement;
    if (frame?.localName === "graphicFrame") {
      const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
      if (nextModern && frame.parentElement?.namespaceURI !== MC) {
        const alternate = drawing.createElementNS(MC, "mc:AlternateContent");
        const choice = drawing.createElementNS(MC, "mc:Choice");
        choice.setAttribute("xmlns:cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
        choice.setAttribute("Requires", "cx1");
        frame.parentNode!.insertBefore(alternate, frame);
        alternate.appendChild(choice);
        choice.appendChild(frame);
        const fallback = drawing.createElementNS(MC, "mc:Fallback");
        const shape = parseWorkbookXml(`<xdr:sp xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="${A}"><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr><xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>This chart requires a newer version of Excel.</a:t></a:r></a:p></xdr:txBody></xdr:sp>`).documentElement;
        fallback.appendChild(drawing.importNode(shape, true));
        alternate.appendChild(fallback);
      } else if (!nextModern && frame.parentElement?.localName === "Choice") {
        const alternate = frame.parentElement.parentElement!;
        alternate.parentNode!.insertBefore(frame, alternate);
        alternate.parentNode!.removeChild(alternate);
      }
    }
    const types = parseWorkbookXml(strFromU8(assets.archive["[Content_Types].xml"]));
    const contentType = Array.from(types.documentElement.children).find(node => node.getAttribute("PartName") === `/${origin.chartPath}`);
    if (!contentType) throw new Error("The chart content type is missing.");
    contentType.setAttribute("ContentType", nextModern ? "application/vnd.ms-office.chartex+xml" : "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
    for (const [path, doc] of [[origin.drawingPath, drawing], [relPath, relationships], ["[Content_Types].xml", types]] as const) {
      assets.archive[path] = strToU8(serializeWorkbookXml(doc));
      assets.dirtyArchivePaths.add(path);
    }
  }
}
