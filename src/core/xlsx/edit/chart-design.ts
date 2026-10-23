// Chart Design > Chart Styles and Quick Layout for an xlsx `ChartObject`, through the shared chart
// catalogue (`chart/style-presets`, `chart/quick-layouts`) and edit operations (`chart/edit`).
//
// Chart Styles: the curated presets every product's gallery offers. A pick pins each series' preset
// colour (`c:ser/c:spPr`), or each point's on a pie or doughnut, as Change Colors pins a palette;
// the series' point overrides are cleared. Excel's own 2013+ chart styles (`styleN.xml` parts, ids
// 201+) and the preset's `c:style` number are not written: the summary carries no style id.
//
// Quick Layout: the title, legend and bar spacing go to the summary fields the chart already saves;
// the axis titles, gridlines, value-axis visibility, data labels and data table are applied to the
// summary for drawing and recorded as `quickLayout`, which the writer applies to the part with
// `applyChartQuickLayout` on save (axis titles and the data table are saved but not drawn).
import type { DrawingColor } from '../../drawingml/types';
import { applyChartQuickLayout, CHART_QUICK_LAYOUT_KINDS } from '../../chart/edit/quick-layout';
import { findChartQuickLayout } from '../../chart/quick-layouts';
import { chartSummaryFromSpace } from '../../chart/render/summary-from-space';
import { chartStylePresetOfColors, findChartStylePreset } from '../../chart/style-presets';
import type { ChartObject, ChartSeries } from '../model';
import { chartSpaceFromObject } from '../write/chart-space';
import type { ChartPatch } from './charts';

const srgb = (hex: string): DrawingColor => ({
	kind: 'srgb',
	value: hex.replace('#', '').toUpperCase(),
	transforms: [],
});

const radial = (chart: ChartObject): boolean =>
	chart.chartType === 'pie' || chart.chartType === 'doughnut';

const hexOf = (color: DrawingColor | undefined): string | undefined =>
	color?.kind === 'srgb' && color.transforms.length === 0 ? `#${color.value}` : undefined;

/** The Chart Styles preset the chart's pinned colours show, if any. */
export function chartStylePresetOf(chart: ChartObject): string | undefined {
	const first = chart.series[0];
	const colors = radial(chart)
		? (first?.values ?? []).map((_, point) => hexOf(first?.pointColors?.[point]))
		: chart.series.map((series) => (series.fill ? undefined : hexOf(series.drawingColor)));
	return chartStylePresetOfColors(colors)?.id;
}

function pinned(series: ChartSeries, color: DrawingColor | undefined): ChartSeries {
	const copy = structuredClone(series);
	delete copy.color;
	delete copy.fill;
	delete copy.pointFills;
	delete copy.pointColors;
	if (color) copy.drawingColor = color;
	return copy;
}

/** The patch applying Chart Styles preset `id`, or `undefined` for an unknown preset. */
export function chartStylePresetPatch(chart: ChartObject, id: string): ChartPatch | undefined {
	const preset = findChartStylePreset(id);
	if (!preset) return undefined;
	const colors = preset.colors();
	const at = (i: number) => srgb(colors[i % colors.length]!);
	if (radial(chart))
		return {
			series: chart.series.map((series) => {
				const copy = pinned(series, undefined);
				copy.pointColors = Object.fromEntries(series.values.map((_, point) => [point, at(point)]));
				return copy;
			}),
		};
	return { series: chart.series.map((series, index) => pinned(series, at(index))) };
}

/** Whether the chart takes Quick Layouts (column, bar, line and area charts). */
export function chartTakesQuickLayout(chart: ChartObject): boolean {
	const kind = chart.chartType === 'column' ? 'bar' : chart.chartType;
	return CHART_QUICK_LAYOUT_KINDS.has(kind);
}

/** The title a layout shows when the chart had none: the series name of a one-series chart. */
const defaultTitle = (chart: ChartObject): string =>
	(chart.series.length === 1 && chart.series[0]?.name) || 'Chart Title';

/** The patch applying Quick Layout `n`, or `undefined` for an unknown layout or chart family. */
export function chartQuickLayoutPatch(chart: ChartObject, n: number): ChartPatch | undefined {
	const layout = findChartQuickLayout(n);
	if (!layout || !chartTakesQuickLayout(chart)) return undefined;
	const { quickLayout: _previous, ...rest } = chart;
	const base = chartSpaceFromObject({
		...rest,
		series: chart.series.map((s) => pinned(s, undefined)),
	});
	if (chart.axes?.length) base.plotArea.axes = structuredClone(chart.axes);
	const next = applyChartQuickLayout(base, n, ['dataLabels', 'axes', 'dataTable']);
	const summary = chartSummaryFromSpace(next);
	const series = chart.series.map((source, index) => {
		const copy = structuredClone(source);
		const labels = summary.series[index]?.dataLabels?.chart;
		if (layout.dataLabels && labels) copy.dataLabels = { ...copy.dataLabels, chart: labels };
		else delete copy.dataLabels;
		return copy;
	});
	const patch: Record<string, unknown> = {
		quickLayout: n,
		title: layout.title ? (chart.title ?? defaultTitle(chart)) : undefined,
		showLegend: layout.legend !== null,
		axes: next.plotArea.axes,
		series,
	};
	if (layout.legend) patch['legendPosition'] = layout.legend;
	if (chart.chartType === 'bar' || chart.chartType === 'column') {
		patch['barGapWidth'] = layout.gapWidth;
		patch['barOverlap'] = layout.overlap;
	}
	return patch as ChartPatch;
}
