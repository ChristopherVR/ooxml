// How the chart edit operations address a `ChartSpace`: every operation is a pure function that
// returns a new model, built by mutating a deep copy (`editChartSpace`). Series are addressed by
// their position in the flattened, document-order list of every chart group's series, which is the
// index PowerPoint's and Excel's editors show (a combination chart's groups follow each other).
import type { ChartPlotGroup, ChartSpace } from '../model';
import type { ChartSpaceSeries } from '../model-series';

/** A copy of `space` changed by `recipe`; `space` itself is never mutated. */
export function editChartSpace(space: ChartSpace, recipe: (draft: ChartSpace) => void): ChartSpace {
	const draft = structuredClone(space);
	recipe(draft);
	return draft;
}

/** One series with the group that holds it. */
export interface ChartSeriesEntry {
	group: ChartPlotGroup;
	groupIndex: number;
	series: ChartSpaceSeries;
	/** Position inside the group's series list. */
	position: number;
}

/** Every series of every group, in document order. */
export function chartSeriesEntries(space: ChartSpace): ChartSeriesEntry[] {
	return space.plotArea.groups.flatMap((group, groupIndex) =>
		group.series.map((series, position) => ({ group, groupIndex, series, position })),
	);
}

/** Every series of every group, in document order. */
export const chartSpaceSeries = (space: ChartSpace): ChartSpaceSeries[] =>
	chartSeriesEntries(space).map((entry) => entry.series);

/**
 * The series at flattened position `index`. The index comes from callers at run time, so it is
 * compared as a position and never used as a property key (a key such as `__proto__` must not
 * select an object to write through).
 */
export function chartSeriesEntryAt(space: ChartSpace, index: number): ChartSeriesEntry {
	const entries = chartSeriesEntries(space);
	const entry =
		typeof index === 'number' && Number.isSafeInteger(index)
			? entries.find((_, position) => position === index)
			: undefined;
	if (!entry) throw new RangeError(`No chart series at index ${String(index)}`);
	return entry;
}

/** The first chart group; edits that need one throw when the plot area has none. */
export function firstChartGroup(space: ChartSpace): ChartPlotGroup {
	const group = space.plotArea.groups[0];
	if (!group) throw new RangeError('The chart has no chart group');
	return group;
}

/** Sets `key` to `value`, or removes it when `value` is `undefined` (exact optional properties). */
export function assignOrDelete<T extends object, K extends keyof T>(
	target: T,
	key: K,
	value: T[K] | undefined,
): void {
	if (value === undefined) delete target[key];
	else target[key] = value;
}
