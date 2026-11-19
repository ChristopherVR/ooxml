/**
 * The workbook writes that bring an embedded workbook in line with an edit of a chart: one per
 * series name, category and value reference (`c:f`), naming only the cells the edit changed. A cell
 * the edit did not touch keeps what it holds, which the cache may not show: a chart caches a text
 * cell in a numeric range as 0, and a blank point can be a cell with anything in it. Literal sources
 * (`c:strLit`, `c:numLit`) and multi-level categories have no single range to write and give none.
 *
 * @module chart/workbook/space-writes
 */
import type { ChartSpace } from '../model';
import type { ChartDataCache, ChartDataSource, ChartSpaceSeries } from '../model-series';
import { parseChartSpace } from '../parse-space';
import { chartPointsLength, slottedPoints, type ChartPointCache } from '../point-limit';
import { parseChartFormulaRange, rangeCellAddress } from './cellref';
import type { ChartWorkbookCell, ChartWorkbookWrite } from './writer';

type TextCache = ChartPointCache<{ index?: number | undefined; value: string }>;

const cellTexts = (cache: TextCache | undefined) =>
	new Map(cache ? slottedPoints(cache).map(([slot, point]) => [slot, point.value]) : []);

/** The cells that make `before` read as `cache`: new or changed points, and emptied ones. */
export function chartCacheCellChanges(
	cache: TextCache,
	before: TextCache | undefined,
): ChartWorkbookCell[] {
	const after = cellTexts(cache);
	const prior = cellTexts(before);
	const cells: ChartWorkbookCell[] = [];
	for (const [index, value] of after) if (prior.get(index) !== value) cells.push({ index, value });
	for (const index of prior.keys()) if (!after.has(index)) cells.push({ index, value: '' });
	return cells;
}

/** The cells a cache's range holds: its declared count, or as many as the cache reaches. */
export const chartCacheCellCount = (cache: TextCache): number =>
	Math.max(cache.pointCount ?? 0, chartPointsLength(cache));

/** Where a reference starts, which stays put when a range grows or shrinks at its end. */
function rangeStart(formula: string): string {
	const range = parseChartFormulaRange(formula);
	return range
		? `${range.sheetName.toLowerCase()}!${range.start.col},${range.start.row}`
		: formula.trim().toLowerCase();
}

type WritableSource = ChartDataSource & { formula: string; cache: ChartDataCache };

/** The sources of a series that name one range and cache its values. */
const writableSources = (series: ChartSpaceSeries): WritableSource[] =>
	[
		series.tx?.reference,
		series.categories,
		series.values,
		series.xValues,
		series.yValues,
		series.bubbleSizes,
	].filter(
		(source): source is WritableSource =>
			!!source?.formula && !!source.cache && source.kind !== 'multiLvlStrRef',
	);

const allSeries = (space: ChartSpace) => space.plotArea.groups.flatMap((group) => group.series);

/**
 * The writes that make the workbook cells hold what `space` caches: every cached point, or with
 * `before` (the chart as loaded) only the cells the edit changed, and for a range `before` did not
 * read, its cached points with every other point-like cell emptied (`only`). A range two series
 * share is written once, from the first; a series added as a copy of another has no cells of its
 * own yet.
 */
export function chartSpaceWorkbookWrites(
	space: ChartSpace,
	before?: ChartSpace,
): ChartWorkbookWrite[] {
	const prior = new Map<string, ChartDataCache>();
	for (const series of before ? allSeries(before) : [])
		for (const source of writableSources(series)) {
			const start = rangeStart(source.formula);
			if (!prior.has(start)) prior.set(start, source.cache);
		}
	const seen = new Set<string>();
	const writes: ChartWorkbookWrite[] = [];
	for (const series of allSeries(space))
		for (const { formula, cache } of writableSources(series)) {
			const start = rangeStart(formula);
			if (seen.has(start)) continue;
			seen.add(start);
			const earlier = prior.get(start);
			const cells = chartCacheCellChanges(cache, earlier);
			// A range `before` did not read may hold anything (a series removed and brought back by
			// undo, data put back on a sheet it had left): the cache is all it holds.
			const only = before !== undefined && earlier === undefined;
			if (cells.length || !before || only)
				writes.push({
					formula,
					isNumeric: cache.type === 'number',
					cells,
					cellCount: chartCacheCellCount(cache),
					...(only ? { only } : {}),
				});
		}
	return writes;
}

/** `formula`'s conflicts in `space`, each under a key that stays put while the range moves its end. */
function rangeConflicts(space: ChartSpace): Map<string, string> {
	const cachedBy = new Map<string, string>();
	const conflicts = new Map<string, string>();
	for (const series of allSeries(space))
		for (const { formula, cache } of writableSources(series)) {
			const range = parseChartFormulaRange(formula);
			if (!range) continue;
			const { start, end } = range;
			const count = chartCacheCellCount(cache);
			const cells = start.col === end.col || start.row === end.row;
			const size = cells ? Math.max(end.col - start.col, end.row - start.row) + 1 : NaN;
			if (!rangeCellAddress(range, count))
				conflicts.set(`${rangeStart(formula)}|${size - count}`, formula);
			const whole = `${range.sheetName.toLowerCase()}!${start.col},${start.row}:${end.col},${end.row}`;
			const points = JSON.stringify([cache.type, ...cellTexts(cache)]);
			const first = cachedBy.get(whole) ?? points;
			cachedBy.set(whole, first);
			if (first !== points) conflicts.set(`${whole}|shared`, formula);
		}
	return conflicts;
}

/**
 * The references of `space` whose range cannot hold what they cache: a range of another size than
 * its cache, or a range that two references cache differently. Writing such a part's caches cell
 * for cell would leave range, cache and workbook disagreeing; its data needs ranges of its own
 * (`rebaseChartData`). With `before` (the parts `space` replaces), only the conflicts none of them
 * had: a chart over hidden rows that shows only the visible ones caches fewer points than its
 * range names, and keeps doing so through its edits. A reference that names no plain range (a
 * defined name, another workbook) is not judged.
 */
export function chartSpaceRangeConflicts(space: ChartSpace, ...before: ChartSpace[]): string[] {
	const had = new Set(before.flatMap((part) => [...rangeConflicts(part).keys()]));
	return [...rangeConflicts(space)].filter(([key]) => !had.has(key)).map(([, formula]) => formula);
}

/** The writes for a chart part saved as `after` from `before`; none when either cannot be read. */
export function chartPartWorkbookWrites(before: string, after: string): ChartWorkbookWrite[] {
	try {
		return chartSpaceWorkbookWrites(
			parseChartSpace(after).chartSpace,
			parseChartSpace(before).chartSpace,
		);
	} catch {
		return [];
	}
}
