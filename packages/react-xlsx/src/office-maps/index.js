import init, { renderChartSvg } from './office_charts.js';
import { createLoader } from './loader.js';
export const loadOfficeCharts = createLoader(init, renderChartSvg, () => import('./payload.js'));
