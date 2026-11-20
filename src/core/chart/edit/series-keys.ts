// The series fields a plot group keeps its data in. Internal to the edit operations.
import type { ChartGroupKind, ChartPlotGroup } from '../model';

/** Scatter and bubble series plot x and y values instead of categories and values. */
export const isXYChartKind = (kind: ChartGroupKind): boolean =>
	kind === 'scatter' || kind === 'bubble';

/** The series field holding a group's values: y values on scatter and bubble charts. */
export const chartValueKey = (group: ChartPlotGroup): 'yValues' | 'values' =>
	isXYChartKind(group.kind) ? 'yValues' : 'values';
/** The series field holding a group's categories: x values on scatter and bubble charts. */
export const chartCategoryKey = (group: ChartPlotGroup): 'xValues' | 'categories' =>
	isXYChartKind(group.kind) ? 'xValues' : 'categories';
