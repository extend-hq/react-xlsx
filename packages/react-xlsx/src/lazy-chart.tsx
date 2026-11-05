import type { OfficeCharts } from "@extend-ai/extend-office-charts";
import * as React from "react";
import { chartRenderer, officeChartRenderer, suppliedOfficeChartRenderer, officeMapRenderer } from "./chart-loader";
import { useChartModule } from "./use-chart-module";
import type { ChartSvgProps } from "./chart-renderer";

class ChartLoadBoundary extends React.Component<{
  children: React.ReactNode;
}, { failed: boolean }> {
  state = { failed: false };

  static getDerivedStateFromError() {
    return { failed: true };
  }

  render() {
    if (this.state.failed) {
      return <div role="status">Unable to load chart.</div>;
    }
    return this.props.children;
  }
}

type LazyChartProps = ChartSvgProps & { experimentalOfficeCharts?: OfficeCharts; loadingFallback: (stage: "renderer" | "map") => React.ReactNode };

function MapChartSvg(props: LazyChartProps) {
  const module = useChartModule(officeMapRenderer);
  if (!module) return <>{props.loadingFallback("map")}</>;
  const fallback = <div role="status">Map data is unavailable.</div>;
  return module.renderer && module.SharedChartSvg && props.chart.renderSource
    ? <module.SharedChartSvg {...props} source={props.chart.renderSource} renderer={module.renderer} fallback={fallback} />
    : fallback;
}

function LegacyChartSvg({ loadingFallback, ...props }: LazyChartProps) {
  const renderer = useChartModule(chartRenderer);
  return renderer ? <renderer.MemoChartSvg {...props} /> : <>{loadingFallback("renderer")}</>;
}

function SharedChart(props: LazyChartProps & { source: NonNullable<ChartSvgProps["chart"]["renderSource"]> }) {
  const module = useChartModule(props.experimentalOfficeCharts ? suppliedOfficeChartRenderer : officeChartRenderer);
  if (!module) return <>{props.loadingFallback("renderer")}</>;
  const fallback = <LegacyChartSvg {...props} />;
  const renderer = props.experimentalOfficeCharts ?? module.renderer;
  return renderer && module.SharedChartSvg
    ? <module.SharedChartSvg {...props} renderer={renderer} fallback={fallback} />
    : fallback;
}

export function MemoChartSvg(props: LazyChartProps) {
  return (
    <ChartLoadBoundary key={props.chart.id}>
      {props.chart.chartType === "RegionMap"
        ? <MapChartSvg {...props} />
        : props.chart.renderSource && (!props.chart.renderSource.experimental || props.experimentalOfficeCharts)
        ? <SharedChart {...props} source={props.chart.renderSource} />
        : <LegacyChartSvg {...props} />}
    </ChartLoadBoundary>
  );
}
