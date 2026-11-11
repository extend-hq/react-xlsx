import type * as React from "react";
import type { XlsxChart, XlsxChartElementSelection, XlsxImageRect, XlsxChartRenderSource } from "./types";

type ChartRendererPalette = {
  border: string;
  mutedText: string;
  surface: string;
  text: string;
};

export type ChartWorkerResult = { svg: string | null; pointMaps?: number[][] } | null;
export type RenderChartAsync = (chart: XlsxChart, source: XlsxChartRenderSource, width: number, height: number) => Promise<ChartWorkerResult>;

export type ChartSvgProps = {
  renderChartAsync?: RenderChartAsync;
  chart: XlsxChart;
  onChartElementDoubleClick?: (selection: XlsxChartElementSelection, event: React.MouseEvent<SVGSVGElement>) => void;
  onChartElementPointerDown?: (selection: XlsxChartElementSelection, event: React.PointerEvent<SVGSVGElement>) => void;
  palette: ChartRendererPalette;
  rect: XlsxImageRect;
  selectedChartElement?: XlsxChartElementSelection | null;
};

