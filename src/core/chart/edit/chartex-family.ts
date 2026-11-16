// ChartEx families as Office names them. A pareto chart is not a layout of its own: it is a
// `clusteredColumn` series (binned, or aggregated by category) with a `paretoLine` series that
// names it through `@ownerIdx`, a category axis without gaps, a value axis and a third value axis
// from 0 to 1 shown as percentages. That is the markup Excel 16 writes for a new Pareto chart
// (`AddChart2(-1, xlPareto)`, observed output); the cumulative line itself is computed by the
// consumer and never cached in the part.
import type { ChartExAxis, ChartExGeography, ChartExSeries, ChartExSpace } from '../ex/model';
import { editChartExSpace } from './chartex';
import { CHART_EX_LAYOUTS, setChartExLayout, type ChartExLayout } from './chartex-structure';

/** Every family the insert and Change Chart Type pickers offer, in Office's gallery order. */
export const CHART_EX_FAMILIES = [
	...CHART_EX_LAYOUTS.slice(0, 2),
	'pareto',
	...CHART_EX_LAYOUTS.slice(2),
] as const;
export type ChartExFamily = ChartExLayout | 'pareto';

/** Whether a series is drawn from another series (a pareto line) rather than its own data. */
export const isChartExOwnedSeries = (series: ChartExSeries): boolean =>
	series.ownerIndex !== undefined;

/** Whether the chart is a pareto chart (a histogram with a `paretoLine` series). */
export function isChartExPareto(space: ChartExSpace): boolean {
	return space.plotArea.series.some(
		(series) => series.layoutId === 'paretoLine' && series.ownerIndex !== undefined,
	);
}

/** The family of the chart: `pareto`, else the layout of its first data series. */
export function chartExFamilyOf(space: ChartExSpace): ChartExFamily | undefined {
	if (isChartExPareto(space)) return 'pareto';
	const layout = space.plotArea.series.find((series) => !isChartExOwnedSeries(series))?.layoutId;
	return (CHART_EX_LAYOUTS as readonly string[]).includes(layout ?? '')
		? (layout as ChartExLayout)
		: undefined;
}

/** The three axes Office gives a new pareto chart. */
function paretoAxes(): ChartExAxis[] {
	return [
		{ id: 0, scaling: { kind: 'cat', gapWidth: 0 }, tickLabels: {} },
		{ id: 1, scaling: { kind: 'val' }, majorGridlines: {}, tickLabels: {} },
		{
			id: 2,
			scaling: { kind: 'val', max: 1, min: 0 },
			units: { unit: 'percentage' },
			tickLabels: {},
		},
	];
}

/**
 * Makes a single-series chart a pareto chart. A histogram keeps its binning; any other family
 * aggregates by category, as Office does for text categories. Formatting of the data series and
 * the chart is kept; the axes become Office's pareto axes.
 */
export function toChartExPareto(space: ChartExSpace): ChartExSpace {
	if (isChartExPareto(space)) return space;
	const base = space.plotArea.series.filter((series) => !isChartExOwnedSeries(series));
	if (base.length !== 1) throw new RangeError('A pareto chart has exactly one data series.');
	const wasHistogram = base[0]!.layoutId === 'clusteredColumn';
	const histogram = setChartExLayout(space, 'clusteredColumn');
	return editChartExSpace(histogram, (draft) => {
		const index = draft.plotArea.series.findIndex((series) => !isChartExOwnedSeries(series));
		const series = draft.plotArea.series[index]!;
		const layout = (series.layoutProperties ??= {});
		const binned =
			wasHistogram &&
			layout.binning !== undefined &&
			(layout.binning.binSize !== undefined || layout.binning.binCount !== undefined);
		if (!binned) {
			delete layout.binning;
			layout.aggregation = true;
		}
		draft.plotArea.axes = paretoAxes();
		series.axisIds = [1];
		draft.plotArea.series.splice(index + 1, 0, {
			layoutId: 'paretoLine',
			ownerIndex: index,
			dataPoints: [],
			axisIds: [2],
		});
	});
}

/** Removes the pareto line and its percentage axis; the histogram stays. */
export function fromChartExPareto(space: ChartExSpace): ChartExSpace {
	if (!isChartExPareto(space)) return space;
	return editChartExSpace(space, (draft) => {
		const owned = draft.plotArea.series.filter(isChartExOwnedSeries);
		const kept = draft.plotArea.series.filter((series) => !isChartExOwnedSeries(series));
		const used = new Set(kept.flatMap((series) => series.axisIds));
		const dropped = new Set(
			owned.flatMap((series) => series.axisIds).filter((id) => !used.has(id)),
		);
		draft.plotArea.series = kept;
		draft.plotArea.axes = draft.plotArea.axes.filter(
			(axis) => axis.id === undefined || !dropped.has(axis.id),
		);
		for (const series of kept)
			series.axisIds = draft.plotArea.axes.flatMap((axis) =>
				axis.id === undefined ? [] : [axis.id],
			);
	});
}

/** Change Chart Type across ChartEx families, pareto included. */
export function setChartExFamily(space: ChartExSpace, family: ChartExFamily): ChartExSpace {
	if (family === 'pareto') return toChartExPareto(space);
	return setChartExLayout(fromChartExPareto(space), family);
}

/**
 * The geography of a new region map: the culture Office resolves place names in, and the
 * projection and area when the author chose them. `attribution` is required by the schema; a part
 * written here has no provider cache, so it names no provider. Office fills `cx:geoCache` itself,
 * through its own online service, when it opens and refreshes the chart.
 */
export function newChartExGeography(
	options: Partial<Omit<ChartExGeography, 'geoCacheXml' | 'extLst'>> = {},
): ChartExGeography {
	return {
		...(options.projectionType ? { projectionType: options.projectionType } : {}),
		...(options.viewedRegionType ? { viewedRegionType: options.viewedRegionType } : {}),
		cultureLanguage: options.cultureLanguage ?? 'en-US',
		cultureRegion: options.cultureRegion ?? 'US',
		attribution: options.attribution ?? '',
	};
}

/** The cell references of a chart's data, as a host knows them (Excel's selection). */
export interface ChartExSourceReferences {
	/** The category range; several columns are hierarchy levels, outermost first. */
	categories?: string;
	/** Per data series, in sheet order: its name cell and values range. */
	series: { name?: string; values?: string }[];
}

/**
 * Points a chart's dimensions and series names at worksheet cells, keeping the cached values. A
 * chart made from literal data becomes one Excel refreshes from its sheet.
 */
export function bindChartExReferences(
	space: ChartExSpace,
	references: ChartExSourceReferences,
): ChartExSpace {
	return editChartExSpace(space, (draft) => {
		const plot = draft.plotArea.series.flatMap((series, index) =>
			isChartExOwnedSeries(series) ? [] : [index],
		);
		plot.forEach((position, row) => {
			const series = draft.plotArea.series[position]!;
			const source = references.series[row];
			if (!source) return;
			const data = series.data ?? draft.chartData.data.find((data) => data.id === series.dataId);
			for (const dimension of data?.dimensions ?? []) {
				const formula = dimension.kind === 'strDim' ? references.categories : source.values;
				if (formula) dimension.formula = { formula };
			}
			if (source.name)
				series.tx = {
					formula: { formula: source.name },
					value: series.tx?.value ?? series.tx?.rich?.text ?? '',
				};
		});
	});
}
