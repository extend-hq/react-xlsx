import type * as React from "react";
import type { XlsxChart, XlsxChartElementSelection, XlsxImageRect } from "./types";

type ChartRendererPalette = {
  border: string;
  mutedText: string;
  surface: string;
  text: string;
};

export type ChartSvgProps = {
  chart: XlsxChart;
  onChartElementDoubleClick?: (selection: XlsxChartElementSelection, event: React.MouseEvent<SVGSVGElement>) => void;
  onChartElementPointerDown?: (selection: XlsxChartElementSelection, event: React.PointerEvent<SVGSVGElement>) => void;
  palette: ChartRendererPalette;
  rect: XlsxImageRect;
  selectedChartElement?: XlsxChartElementSelection | null;
};

