// The value axis of a cartesian chart view: percent, stacked or plain extents, nicely scaled.
import { niceScale, PERCENT_SCALE } from './chart-scale';
import type { ChartViewModel, ValueAxisView } from './chart-view';

/** The value (Y) axis scale for a cartesian view, from its grouping and series values. */
export function valueAxis(model: ChartViewModel): ValueAxisView {
	const stackable =
		model.type === 'bar' ||
		model.type === 'column' ||
		model.type === 'line' ||
		model.type === 'area';
	if (stackable && model.grouping === 'percentStacked')
		return { ...PERCENT_SCALE, ticks: [...PERCENT_SCALE.ticks], percent: true };
	let min = Infinity;
	let max = -Infinity;
	if (stackable && model.grouping === 'stacked') {
		for (let i = 0; i < model.categories.length; i++) {
			let pos = 0;
			let neg = 0;
			for (const s of model.series) {
				const v = s.values[i] ?? 0;
				if (v >= 0) pos += v;
				else neg += v;
			}
			min = Math.min(min, neg);
			max = Math.max(max, pos);
		}
	} else {
		for (const s of model.series)
			for (const v of s.values)
				if (v !== null) {
					min = Math.min(min, v);
					max = Math.max(max, v);
				}
	}
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
