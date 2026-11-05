import type { OfficeCharts } from "@extend-ai/extend-office-charts";
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
  try {
    const [{ loadOfficeCharts }, { SharedChartSvg }] = await Promise.all([
      import("@extend-ai/extend-office-charts"), import("./shared-chart")
    ]);
    return { renderer: await loadOfficeCharts(), SharedChartSvg };
  } catch {
    return { renderer: null, SharedChartSvg: null };
  }
});

export const suppliedOfficeChartRenderer = createChartModule<OfficeChartModule>(async () => {
  const { SharedChartSvg } = await import("./shared-chart");
  return { renderer: null, SharedChartSvg };
});

export const chartRenderer = createChartModule(() => import("./chart-renderer"));
export const officeMapRenderer = createChartModule<OfficeChartModule>(async () => {
  const [{ loadOfficeCharts }, { SharedChartSvg }] = await Promise.all([
    import("./office-maps/index.js"), import("./shared-chart")
  ]);
  return { renderer: await loadOfficeCharts(), SharedChartSvg };
});
export const surfaceRenderer = createChartModule(() => import("./surface-webgl"));

export function preloadChartRenderers(chartsBySheet: readonly (readonly XlsxChart[])[], experimentalOfficeCharts?: OfficeCharts) {
  let hasCharts = false;
  let hasSharedCharts = false;
  let hasExtendedCharts = false;
  let hasMaps = false;
  let hasSurfaces = false;
  for (const charts of chartsBySheet) {
    for (const chart of charts) {
      if (chart.chartType === "RegionMap") {
        hasMaps = true;
        continue;
      }
      const source = chart.renderSource;
      hasSharedCharts ||= Boolean(source && !source.experimental);
      hasExtendedCharts ||= Boolean(source?.experimental && experimentalOfficeCharts);
      const fallback = !source || Boolean(source.experimental && !experimentalOfficeCharts);
      hasCharts ||= fallback;
      hasSurfaces ||= fallback && chart.chartType === "Surface";
    }
  }

  if (experimentalOfficeCharts && (hasExtendedCharts || hasSharedCharts)) void suppliedOfficeChartRenderer.load().catch(() => {});
  else if (hasSharedCharts) void officeChartRenderer.load().catch(() => {});
  if (hasCharts) void chartRenderer.load().catch(() => {});
  if (hasMaps) void officeMapRenderer.load().catch(() => {});
  if (hasSurfaces) void surfaceRenderer.load().catch(() => {});
}
