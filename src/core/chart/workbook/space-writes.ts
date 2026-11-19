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
import { parseChartFormulaRange } from './cellref';
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
 * `before` (the chart as loaded) only the cells the edit changed. A range two series share is
 * written once, from the first; a series added as a copy of another has no cells of its own yet.
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
			const cells = chartCacheCellChanges(cache, prior.get(start));
			if (cells.length || !before)
				writes.push({
					formula,
					isNumeric: cache.type === 'number',
					cells,
					cellCount: chartCacheCellCount(cache),
				});
		}
	return writes;
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
