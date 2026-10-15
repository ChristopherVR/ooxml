// The value axis of a cartesian chart view: percent, stacked or plain extents, nicely scaled.
import { niceScale, PERCENT_SCALE } from './chart-scale';
import type { ChartSeriesView, ChartViewModel, ValueAxisView } from './chart-view';

/** The smallest and largest value a group draws: stacked groups sum per category. */
export function valueExtent(
	type: string,
	grouping: string,
	series: readonly Pick<ChartSeriesView, 'values'>[],
	categories: number,
): { min: number; max: number } {
	const stackable = type === 'bar' || type === 'column' || type === 'line' || type === 'area';
	let min = Infinity;
	let max = -Infinity;
	if (stackable && grouping === 'stacked') {
		for (let i = 0; i < categories; i++) {
			let pos = 0;
			let neg = 0;
			for (const s of series) {
				const v = s.values[i] ?? 0;
				if (v >= 0) pos += v;
				else neg += v;
			}
			min = Math.min(min, neg);
			max = Math.max(max, pos);
		}
	} else {
		for (const s of series)
			for (const v of s.values)
				if (v !== null) {
					min = Math.min(min, v);
					max = Math.max(max, v);
				}
	}
	return { min, max };
}

/** The value (Y) axis scale for a cartesian view, from its grouping and series values. */
export function valueAxis(model: ChartViewModel): ValueAxisView {
	const stackable =
		model.type === 'bar' ||
		model.type === 'column' ||
		model.type === 'line' ||
		model.type === 'area';
	if (stackable && model.grouping === 'percentStacked')
		return { ...PERCENT_SCALE, ticks: [...PERCENT_SCALE.ticks], percent: true };
	let { min, max } = valueExtent(model.type, model.grouping, model.series, model.categories.length);
	if (min === Infinity) {
		min = 0;
		max = 1;
	}
	return { ...niceScale(min, max), percent: false };
}

/** Per-category stacked totals of absolute values (for percent-stacked charts). */
export function categoryTotals(model: ChartViewModel): number[] {
	return model.categories.map((_, i) =>
		model.series.reduce((sum, s) => sum + Math.abs(s.values[i] ?? 0), 0),
	);
}
