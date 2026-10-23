// Chart data edits of an extended chart (`ChartExSpace`): series values and categories. They follow
// PowerPoint's ChartEx save (pptx `runtime/chart-cx-update-data.ts`): the `cx:data` a series draws
// keeps its formulas, extra dimensions and level attributes (format code, name) and gets fresh
// cached levels; a series that draws no data gets a new `cx:data` with the next free id. Blank
// values (`null`) are written as missing points, as Office writes an empty cell.
import type { ChartSeriesValue } from '../data-sheet';
import type { ChartExData, ChartExDimension, ChartExLevel, ChartExSpace } from '../ex/model';
import { chartExSeriesAt, editChartExSpace } from './chartex';

/** The numeric dimension type a layout draws: `size` (treemap, sunburst), `colorVal`, `val`. */
export function chartExValueType(layoutId: string): string {
	return layoutId === 'treemap' || layoutId === 'sunburst'
		? 'size'
		: layoutId === 'regionMap'
			? 'colorVal'
			: 'val';
}

function freshLevel(
	values: readonly (string | number | null)[],
	previous: ChartExLevel | undefined,
	numeric: boolean,
): ChartExLevel {
	const level: ChartExLevel = {
		pointCount: values.length,
		points: values.flatMap((value, index) =>
			value === null || value === '' ? [] : [{ index, value: String(value) }],
		),
	};
	if (numeric) level.formatCode = previous?.formatCode ?? 'General';
	if (previous?.name !== undefined) level.name = previous.name;
	return level;
}

/** The `cx:data` the series at `index` draws in `draft`, created (and bound) when it has none. */
function dataOf(draft: ChartExSpace, index: number): ChartExData {
	const series = chartExSeriesAt(draft, index);
	if (series.data) return series.data;
	const shared =
		series.dataId === undefined
			? undefined
			: draft.chartData.data.find((data) => data.id === series.dataId);
	if (shared) return shared;
	const id = Math.max(-1, ...draft.chartData.data.map((data) => data.id ?? -1)) + 1;
	const data: ChartExData = { id, dimensions: [] };
	draft.chartData.data.push(data);
	series.dataId = id;
	return data;
}

/** The dimension of `kind` and `type` (PowerPoint's fallbacks: any numeric, an untyped text one). */
function dimensionOf(data: ChartExData, kind: ChartExDimension['kind'], type: string) {
	const all = data.dimensions.filter((d) => d.kind === kind);
	let match = all.find((d) => d.type === type);
	if (!match && kind === 'numDim') match = all[0];
	if (!match && kind === 'strDim' && type === 'cat') match = all.find((d) => d.type === undefined);
	if (!match) {
		match = { kind, type, levels: [] };
		data.dimensions.push(match);
	}
	return match;
}

/**
 * Sets the values of the series at `index` (numbers, or `null` for a blank). Semantics: PowerPoint.
 */
export function setChartExSeriesValues(
	space: ChartExSpace,
	index: number,
	values: readonly ChartSeriesValue[],
): ChartExSpace {
	for (const value of values)
		if (value !== null && !Number.isFinite(value))
			throw new RangeError('Chart values must be finite numbers or null');
	return editChartExSpace(space, (draft) => {
		const layoutId = chartExSeriesAt(draft, index).layoutId;
		const dimension = dimensionOf(dataOf(draft, index), 'numDim', chartExValueType(layoutId));
		dimension.levels = [freshLevel(values, dimension.levels[0], true)];
	});
}

/**
 * Sets the categories of every series: one list, or the levels of a hierarchy (treemap, sunburst;
 * leaf level first). Semantics: PowerPoint (categories belong to the chart, so each series' data
 * gets them).
 */
export function setChartExCategories(
	space: ChartExSpace,
	categories: readonly (readonly string[])[] | readonly string[],
): ChartExSpace {
	const levels = (
		categories.length > 0 && Array.isArray(categories[0]) ? categories : [categories]
	) as readonly (readonly string[])[];
	return editChartExSpace(space, (draft) => {
		const seen = new Set<ChartExData>();
		draft.plotArea.series.forEach((_, index) => {
			const data = dataOf(draft, index);
			if (seen.has(data)) return;
			seen.add(data);
			const dimension = dimensionOf(data, 'strDim', 'cat');
			dimension.levels = levels.map((level, position) =>
				freshLevel(level, dimension.levels[position], false),
			);
		});
	});
}
