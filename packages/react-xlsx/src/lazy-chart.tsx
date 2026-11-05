import type { OfficeCharts } from "@extend-ai/rust-office-charts";
import * as React from "react";
import { chartRenderer, officeChartRenderer, suppliedOfficeChartRenderer, officeExtendedChartRenderer } from "./chart-loader";
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

function LegacyChartSvg({ loadingFallback, ...props }: LazyChartProps) {
  const renderer = useChartModule(chartRenderer);
  return renderer ? <renderer.MemoChartSvg {...props} /> : <>{loadingFallback("renderer")}</>;
}

function SharedChart(props: LazyChartProps & { source: NonNullable<ChartSvgProps["chart"]["renderSource"]> }) {
  const module = useChartModule(props.experimentalOfficeCharts ? suppliedOfficeChartRenderer : props.source.experimental ? officeExtendedChartRenderer : officeChartRenderer);
  if (!module) return <>{props.loadingFallback(props.chart.chartType === "RegionMap" ? "map" : "renderer")}</>;
  const fallback = props.chart.chartType === "RegionMap" ? <div role="status">Map data is unavailable.</div> : <LegacyChartSvg {...props} />;
  const renderer = props.experimentalOfficeCharts ?? module.renderer;
  return renderer && module.SharedChartSvg
    ? <module.SharedChartSvg {...props} renderer={renderer} fallback={fallback} />
    : fallback;
}

export function MemoChartSvg(props: LazyChartProps) {
  return (
    <ChartLoadBoundary key={props.chart.id}>
      {props.chart.renderSource
        ? <SharedChart {...props} source={props.chart.renderSource} />
        : <LegacyChartSvg {...props} />}
    </ChartLoadBoundary>
  );
}
