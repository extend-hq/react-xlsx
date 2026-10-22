import * as React from "react";
import type { ChartModule } from "./chart-loader";

export function useChartModule<T>(resource: ChartModule<T>) {
  const snapshot = React.useSyncExternalStore(resource.subscribe, resource.getSnapshot, resource.getServerSnapshot);
  React.useEffect(() => {
    void resource.load().catch(() => {});
  }, [resource]);
  if (snapshot.error) throw snapshot.error;
  return snapshot.module;
}
