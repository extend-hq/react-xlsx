import * as React from "react";
import { chartRenderer, regionMapResources } from "./chart-loader";
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

function ChartSvg({ loadingFallback, ...props }: LazyChartProps) {
  const renderer = useChartModule(chartRenderer);
  if (!renderer) return <>{loadingFallback("renderer")}</>;
  return props.chart.chartType === "RegionMap"
    ? <MapChartSvg {...props} renderer={renderer} loadingFallback={loadingFallback} />
    : <renderer.MemoChartSvg {...props} />;
}

export function MemoChartSvg(props: LazyChartProps) {
  return (
    <ChartLoadBoundary key={props.chart.id}>
      <ChartSvg {...props} />
    </ChartLoadBoundary>
  );
}
