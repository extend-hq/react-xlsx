import * as React from "react";
import { chartRenderer, officeChartRenderer, regionMapResources } from "./chart-loader";
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

type LazyChartProps = ChartSvgProps & { loadingFallback: (stage: "renderer" | "map") => React.ReactNode };

function MapChartSvg({ renderer, loadingFallback, ...props }: LazyChartProps & {
  renderer: typeof import("./chart-renderer");
}) {
  const resources = useChartModule(regionMapResources);
  return resources ? <renderer.MemoChartSvg {...props} /> : <>{loadingFallback("map")}</>;
}

function LegacyChartSvg({ loadingFallback, ...props }: LazyChartProps) {
  const renderer = useChartModule(chartRenderer);
  if (!renderer) return <>{loadingFallback("renderer")}</>;
  return props.chart.chartType === "RegionMap"
    ? <MapChartSvg {...props} renderer={renderer} loadingFallback={loadingFallback} />
    : <renderer.MemoChartSvg {...props} />;
}

function SharedChart(props: LazyChartProps & { source: NonNullable<ChartSvgProps["chart"]["renderSource"]> }) {
  const module = useChartModule(officeChartRenderer);
  if (!module) return <>{props.loadingFallback("renderer")}</>;
  const fallback = <LegacyChartSvg {...props} />;
  return module.renderer
    ? <module.SharedChartSvg {...props} renderer={module.renderer} fallback={fallback} />
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
