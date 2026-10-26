// PowerPoint view edits mapped onto shared ChartEx operations. Geography cache changes stay unclaimed.
import {
	addChartExSeries,
	removeChartExSeries,
	setChartExLayout,
	setChartExSeriesFill,
	setChartExSeriesValues,
	setChartExWaterfall,
	setChartExBinning,
	setChartExParentLabelLayout,
	setChartExBoxWhisker,
	type ChartExLayout,
} from '../../../chart';
import type { PptxChartData } from '../types';
import type { ChartExEdit } from './chart-ex-neutral-save';
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);
const layouts: Partial<Record<PptxChartData['chartType'], ChartExLayout>> = {
	waterfall: 'waterfall',
	funnel: 'funnel',
	treemap: 'treemap',
	sunburst: 'sunburst',
	boxWhisker: 'boxWhisker',
	histogram: 'clusteredColumn',
	regionMap: 'regionMap',
};
export function planChartExFamilyEdits(
	before: PptxChartData,
	probe: PptxChartData,
	edits: ChartExEdit[],
) {
	if (before.chartType !== probe.chartType) {
		const layout = layouts[probe.chartType];
		if (!layout || layout === 'regionMap') return;
		edits.push((space) => setChartExLayout(space, layout));
		probe.chartType = before.chartType;
	}
	probe.series.forEach((series, i) => {
		const previous = before.series[i];
		if (!previous) return;
		if (series.color !== previous.color) {
			const color = series.color?.replace(/^#/u, '');
			if (color && !/^[a-f\d]{6}$/iu.test(color)) return;
			edits.push((space) =>
				setChartExSeriesFill(
					space,
					i,
					color
						? { kind: 'solid', color: { kind: 'srgb', value: color.toUpperCase(), transforms: [] } }
						: undefined,
				),
			);
			series.color = previous.color;
		}
		if (!same(series.waterfallOptions, previous.waterfallOptions) && series.waterfallOptions) {
			const value = series.waterfallOptions;
			edits.push((space) =>
				setChartExWaterfall(space, i, {
					subtotals: value.subtotalIndices,
					connectorLines: value.connectorLines,
				}),
			);
			series.waterfallOptions = previous.waterfallOptions;
		}
		if (!same(series.histogramOptions, previous.histogramOptions) && series.histogramOptions) {
			const { layout, aggregateByCategory, ...binning } = series.histogramOptions;
			// A Pareto line owns a distinct linked series and cannot be represented as a columns-only switch.
			if (layout === 'pareto' && layout !== previous.histogramOptions?.layout) return;
			edits.push((space) =>
				setChartExBinning(space, i, aggregateByCategory ? 'category' : binning),
			);
			series.histogramOptions = previous.histogramOptions;
		}
		if (
			!same(series.treemapOptions, previous.treemapOptions) &&
			series.treemapOptions?.parentLabelLayout
		) {
			const value = series.treemapOptions.parentLabelLayout;
			edits.push((space) => setChartExParentLabelLayout(space, i, value));
			series.treemapOptions = previous.treemapOptions;
		}
		if (!same(series.boxWhiskerOptions, previous.boxWhiskerOptions) && series.boxWhiskerOptions) {
			const value = series.boxWhiskerOptions;
			edits.push((space) =>
				setChartExBoxWhisker(space, i, {
					quartileMethod: value.quartileMethod,
					meanLine: value.showMeanLine,
					meanMarker: value.showMeanMarker,
					nonoutliers: value.showInnerPoints,
					outliers: value.showOutlierPoints,
				}),
			);
			series.boxWhiskerOptions = previous.boxWhiskerOptions;
		}
	});
}
/** Positional series edits retain the original series metadata; additions clone the shared template. */
export function planChartExSeriesStructure(
	before: PptxChartData,
	probe: PptxChartData,
	edits: ChartExEdit[],
): boolean {
	if (probe.series.length === before.series.length) return true;
	if (!['histogram', 'boxWhisker'].includes(before.chartType) || !probe.series.length) return false;
	const extras = probe.series.slice(before.series.length);
	const allowed = new Set(['name', 'values', 'blanks', 'color']);
	if (extras.some((series) => Object.keys(series).some((key) => !allowed.has(key)))) return false;
	for (let i = before.series.length - 1; i >= probe.series.length; i--)
		edits.push((space) => removeChartExSeries(space, i));
	extras.forEach((series, offset) => {
		const index = before.series.length + offset;
		edits.push((space) => addChartExSeries(space, series.name));
		edits.push((space) =>
			setChartExSeriesValues(
				space,
				index,
				series.values.map((value, point) => (series.blanks?.[point] ? null : value)),
			),
		);
		if (series.color) {
			const color = series.color.replace(/^#/u, '');
			if (!/^[a-f\d]{6}$/iu.test(color)) throw new Error('Invalid ChartEx series color.');
			edits.push((space) =>
				setChartExSeriesFill(space, index, {
					kind: 'solid',
					color: { kind: 'srgb', value: color.toUpperCase(), transforms: [] },
				}),
			);
		}
	});
	return true;
}
