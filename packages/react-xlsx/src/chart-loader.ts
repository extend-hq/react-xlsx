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

export const chartRenderer = createChartModule(() => import("./chart-renderer"));
export const regionMapResources = createChartModule(() => import("./region-map-geography"));
export const surfaceRenderer = createChartModule(() => import("./surface-regl"));

export function preloadChartRenderers(chartsBySheet: readonly (readonly XlsxChart[])[]) {
  let hasCharts = false;
  let hasMaps = false;
  let hasSurfaces = false;
  for (const charts of chartsBySheet) {
    for (const chart of charts) {
      hasCharts = true;
      hasMaps ||= chart.chartType === "RegionMap";
      hasSurfaces ||= chart.chartType === "Surface";
    }
  }

  if (hasCharts) void chartRenderer.load().catch(() => {});
  if (hasMaps) void regionMapResources.load().catch(() => {});
  if (hasSurfaces) void surfaceRenderer.load().catch(() => {});
}
