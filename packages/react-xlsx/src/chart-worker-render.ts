import { projectChartData } from "./chart-data-projection";
import type { ChartData, OfficeCharts } from "./office-chart-types";
import type { XlsxChart, XlsxChartRenderSource } from "./types";
export function renderChartInWorker(renderer: OfficeCharts, chart: XlsxChart, source: XlsxChartRenderSource, width: number, height: number) {
  try {
    const data: ChartData = {
      title: chart.title,
      series: (source.sourceSeries ?? chart.series).map((series, ordinal) => {
        const chartType = source.seriesTypes?.[ordinal] ?? source.kind;
        const xy = chartType === "scatter" || chartType === "bubble";
        const indexedX = source.indexedXSeries?.[ordinal] === true;
        const numericCategories = series.raw?.chartCategoryNumericValues as Array<number | null> | undefined;
        const useNumericCategories = source.dateCategorySeries?.[ordinal] === true || (xy && !indexedX);
        const hierarchy = series.raw?.chartExHierarchyCategories as string[][] | undefined;
        const colorStrings = series.raw?.chartExColorStrings as Array<string | null> | undefined;
        const classicHierarchy = series.raw?.chartHierarchyCategories as Array<Array<{ level: number; label: string }>> | undefined;
        return {
          chartType,
          name: series.name,
          points: Array.from({ length: Math.max(series.values.length, indexedX ? 0 : series.categories.length, series.bubbleSizes?.length ?? 0) }, (_, index) => {
            const value = series.values[index];
            const label = useNumericCategories && numericCategories?.[index] != null ? numericCategories[index] : series.categories[index];
            const bubble = series.bubbleSizes?.[index];
            return {
              index,
              value: xy ? undefined : colorStrings?.[index] ?? (value == null ? null : String(value)),
              x: indexedX ? String(index + 1) : xy && label != null ? String(label) : undefined,
              y: xy && value != null ? String(value) : undefined,
              bubbleSize: bubble == null ? undefined : String(bubble),
              categories: xy ? [] : hierarchy?.[index]?.length
                ? [...hierarchy[index]].reverse().map((label, level) => ({level, label}))
                : classicHierarchy?.[index]?.length ? classicHierarchy[index]
                : series.categories.length || series.categoriesRef?.formula
                  ? [{ level: 0, label: label == null ? "" : String(label) }]
                  : []
            };
          })
        };
      })
    };
    const projected = projectChartData(source.xml, data, (source.sourceSeries ?? chart.series)
      .map(series => series.raw?.chartExcludedPointIndices as number[] | undefined));
    const svg = renderer.renderChart(projected.xml, {
      themeColors: source.themeColors,
      themeXml: source.themeXml,
      themeOverrideXml: source.themeOverrideXml,
      styleXml: source.styleXml,
      colorStyleXml: source.colorStyleXml,
      images: source.images,
      drawings: source.drawings,
      width, height,
      interactive: true,
      spreadsheetDefaults: true,
      fontFamily: chart.fontFamily,
      data: projected.data
    });
    return { svg, pointMaps: projected.pointMaps };
  } catch {
    return null;
  }
}

