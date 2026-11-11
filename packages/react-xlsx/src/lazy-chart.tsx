import { useXlsxMessages } from "./messages-context";
import type { OfficeCharts } from "@extend-ai/rust-office-charts";
import * as React from "react";
import { officeChartRenderer, suppliedOfficeChartRenderer, officeExtendedChartRenderer } from "./chart-loader";
import { useChartModule } from "./use-chart-module";
import type { ChartSvgProps } from "./chart-renderer-types";

class ChartLoadBoundary extends React.Component<{
  children: React.ReactNode;
  errorMessage: string;
}, { failed: boolean }> {
  state = { failed: false };

  static getDerivedStateFromError() {
    return { failed: true };
  }

  render() {
    if (this.state.failed) {
      return <div role="status">{this.props.errorMessage}</div>;
    }
    return this.props.children;
  }
}

type LazyChartProps = ChartSvgProps & { experimentalOfficeCharts?: OfficeCharts; loadingFallback: (stage: "renderer" | "map") => React.ReactNode };

function UnsupportedChart({ chart }: Pick<ChartSvgProps, "chart">) {
  const messages = useXlsxMessages();
  return <div role="status" data-xlsx-chart-fallback="unsupported">{chart.chartType === "RegionMap"
    ? messages.mapUnavailable
    : messages.chartUnavailable}</div>;
}

function SharedChart(props: LazyChartProps & { source: NonNullable<ChartSvgProps["chart"]["renderSource"]> }) {
  const module = useChartModule(props.experimentalOfficeCharts ? suppliedOfficeChartRenderer : props.source.experimental ? officeExtendedChartRenderer : officeChartRenderer);
  if (!module) return <>{props.loadingFallback(props.chart.chartType === "RegionMap" ? "map" : "renderer")}</>;
  const fallback = <UnsupportedChart chart={props.chart} />;
  const renderer = props.experimentalOfficeCharts ?? module.renderer;
  return module.SharedChartSvg
    ? <module.SharedChartSvg {...props} renderer={renderer} fallback={fallback} />
    : fallback;
}

export function MemoChartSvg(props: LazyChartProps) {
  const messages = useXlsxMessages();
  return (
    <ChartLoadBoundary key={props.chart.id} errorMessage={messages.loadChartFailed}>
      {props.chart.renderSource
        ? <SharedChart {...props} source={props.chart.renderSource} />
        : <UnsupportedChart chart={props.chart} />}
    </ChartLoadBoundary>
  );
}
