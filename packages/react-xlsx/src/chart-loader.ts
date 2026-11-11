import type { OfficeCharts } from "@extend-ai/rust-office-charts";
import type { XlsxChart } from "./types";

type ModuleSnapshot<T> = { module: T | null; error: unknown };

export type ChartModule<T> = {
  load: () => Promise<T>;
  getSnapshot: () => ModuleSnapshot<T>;
  getServerSnapshot: () => ModuleSnapshot<T>;
  subscribe: (listener: () => void) => () => void;
};

function createChartModule<T>(importModule: () => Promise<T>): ChartModule<T> {
  const initialSnapshot: ModuleSnapshot<T> = { module: null, error: null };
  let snapshot = initialSnapshot;
  let promise: Promise<T> | undefined;
  const listeners = new Set<() => void>();
  const publish = (next: ModuleSnapshot<T>) => {
    snapshot = next;
    listeners.forEach((listener) => listener());
  };

  return {
    load: () => promise ??= importModule().then((module) => {
      publish({ module, error: null });
      return module;
    }, (error: unknown) => {
      publish({ module: null, error: error instanceof Error ? error : new Error("Unable to load chart module.") });
      throw error;
    }),
    getSnapshot: () => snapshot,
    getServerSnapshot: () => initialSnapshot,
    subscribe: (listener) => {
      listeners.add(listener);
      return () => { listeners.delete(listener); };
    }
  };
}

type OfficeChartModule = {
  renderer: OfficeCharts | null;
  SharedChartSvg: typeof import("./shared-chart").SharedChartSvg | null;
};

export const officeChartRenderer = createChartModule<OfficeChartModule>(async () => {
  const { SharedChartSvg } = await import("./shared-chart");
  return { renderer: null, SharedChartSvg };
});
export const suppliedOfficeChartRenderer = officeChartRenderer;
export const officeExtendedChartRenderer = officeChartRenderer;

export function preloadChartRenderers(chartsBySheet: readonly (readonly XlsxChart[])[], experimentalOfficeCharts?: OfficeCharts) {
  let hasSharedCharts = false;
  let hasExtendedCharts = false;

  for (const charts of chartsBySheet) {
    for (const chart of charts) {
      const source = chart.renderSource;
      hasSharedCharts ||= Boolean(source && !source.experimental);
      hasExtendedCharts ||= Boolean(source?.experimental);
    }
  }

  if (experimentalOfficeCharts && (hasExtendedCharts || hasSharedCharts)) void suppliedOfficeChartRenderer.load().catch(() => {});
  else if (hasSharedCharts) void officeChartRenderer.load().catch(() => {});
  if (hasExtendedCharts && !experimentalOfficeCharts) void officeExtendedChartRenderer.load().catch(() => {});
}
