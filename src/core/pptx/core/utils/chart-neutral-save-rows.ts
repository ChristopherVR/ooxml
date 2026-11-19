// Category rows added to or removed from a chart between two saves. They go through the same row
// operations as every other product, so a chart whose view shows only part of its declared range
// keeps the rest; what is left of the difference is ordinary cell edits for the planners after this.
import { addChartCategory, removeChartCategory } from '../../../chart/index';
import type { PptxChartData, PptxChartSeries } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';

const rowOf = (chart: PptxChartData, row: number) =>
	JSON.stringify([
		chart.categories[row],
		...chart.series.map((series) => (series.blanks?.[row] ? null : (series.values[row] ?? null))),
	]);

/** The rows of `before` that `after` no longer has, found by the rows either side of each gap. */
function removedRows(before: PptxChartData, after: PptxChartData): number[] {
	const total = before.categories.length;
	const remaining = after.categories.length;
	const removed: number[] = [];
	let left = total - remaining;
	/** How many rows of `before` from `row` are gone, given `after` continues at `at`. */
	const gap = (row: number, at: number) => {
		if (at >= remaining) return left;
		const wanted = rowOf(after, at);
		if (rowOf(before, row) === wanted) return 0;
		for (let ahead = 1; ahead <= left && row + ahead < total; ahead++)
			if (rowOf(before, row + ahead) === wanted) return ahead;
		return 0;
	};
	let at = 0;
	for (let row = 0; row < total;) {
		const skip = left > 0 ? gap(row, at) : 0;
		for (let n = 0; n < skip; n++) removed.push(row + n);
		left -= skip;
		row += skip || 1;
		if (!skip) at += 1;
	}
	return removed;
}

function withoutRows(series: PptxChartSeries, removed: readonly number[]): PptxChartSeries {
	const gone = new Set(removed);
	const kept = <T>(items: readonly T[]) => items.filter((_, row) => !gone.has(row));
	return {
		...series,
		values: kept(series.values),
		...(series.blanks ? { blanks: kept(series.blanks) } : {}),
		...(series.dataPoints
			? {
					dataPoints: series.dataPoints
						.filter((point) => !gone.has(point.idx))
						.map((point) => ({
							...point,
							idx: point.idx - removed.filter((row) => row < point.idx).length,
						})),
				}
			: {}),
	};
}

/**
 * Plans the rows `after` gained or lost and returns `before` as those edits leave it. A chart whose
 * categories are not one plain list keeps to the whole-vector planners.
 */
export function planNeutralCategoryRows(
	before: PptxChartData,
	after: PptxChartData,
	edits: NeutralChartEdit[],
): PptxChartData {
	const added = after.categories.length - before.categories.length;
	if (
		added === 0 ||
		before.series.length !== after.series.length ||
		before.categoryLevels ||
		before.dateCategories ||
		after.dateCategories ||
		['scatter', 'bubble'].includes(before.chartType) ||
		before.series.some((s) => s.seriesChartType === 'scatter' || s.seriesChartType === 'bubble')
	)
		return before;
	if (added > 0) {
		const names = after.categories.slice(before.categories.length);
		edits.push((space) => names.reduce((next, name) => addChartCategory(next, name), space));
		return {
			...before,
			categories: [...before.categories, ...names],
			series: before.series.map((series) => ({
				...series,
				values: [...series.values, ...names.map(() => 0)],
			})),
		};
	}
	const removed = removedRows(before, after);
	edits.push((space) => removed.reduceRight((next, row) => removeChartCategory(next, row), space));
	const gone = new Set(removed);
	return {
		...before,
		categories: before.categories.filter((_, row) => !gone.has(row)),
		series: before.series.map((series) => withoutRows(series, removed)),
	};
}
