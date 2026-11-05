export interface ChartPoint {
  index: number;
  categories?: ReadonlyArray<{ level: number; label: string }>;
  value?: string | null;
  x?: string | null;
  y?: string | null;
  bubbleSize?: string | null;
}

export interface ChartData {
  title?: string | null;
  series: ReadonlyArray<{
    chartType: string;
    name?: string | null;
    points: ReadonlyArray<ChartPoint>;
  }>;
}

export interface RenderOptions {
  width?: number;
  height?: number;
  themeColors?: Readonly<Record<string, string>>;
  images?: Readonly<Record<string, string>>;
  fontFamily?: string;
  themeXml?: string;
  themeOverrideXml?: string;
  styleXml?: string;
  colorStyleXml?: string;
  legacyDefaults?: boolean;
  spreadsheetDefaults?: boolean;
  interactive?: boolean;
  data?: ChartData;
}

export interface OfficeCharts {
  renderChart(xml: string, options?: RenderOptions): string | null;
}

