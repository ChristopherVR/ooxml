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

/**
 * The cells that make `before` read as `cache` when surviving rows have shifted:
 * only points that differ from the surviving row's original value are emitted.
 */
export function chartShiftedCacheCellChanges(
	cache: TextCache,
	before: TextCache | undefined,
	survivingAfterToBefore: ReadonlyMap<number, number>,
): ChartWorkbookCell[] {
	const after = cellTexts(cache);
	const prior = cellTexts(before);
	const cells: ChartWorkbookCell[] = [];
	for (const [index, value] of after) {
		const priorIndex = survivingAfterToBefore.get(index);
		const priorValue = priorIndex !== undefined ? prior.get(priorIndex) : undefined;
		if (priorValue !== value) cells.push({ index, value });
	}
	return cells;
}

function maxRowCount(seriesList: readonly ChartSpaceSeries[]): number {
	let max = 0;
	for (const s of seriesList) {
		const c = s.categories ?? s.xValues;
		if (c?.cache) max = Math.max(max, chartCacheCellCount(c.cache));
		const v = s.values ?? s.yValues;
		if (v?.cache) max = Math.max(max, chartCacheCellCount(v.cache));
	}
	return max;
}

function alignRows(
	space: ChartSpace,
	before: ChartSpace,
):
	| {
			removedRows: number[];
			totalRowsBefore: number;
			survivingAfterToBefore: Map<number, number>;
	  }
	| undefined {
	const seriesBefore = allSeries(before);
	const seriesAfter = allSeries(space);
	const totalBefore = maxRowCount(seriesBefore);
	const totalAfter = maxRowCount(seriesAfter);

	if (totalBefore <= 0 || totalAfter >= totalBefore) return undefined;

	const firstBefore = seriesBefore[0];
	const firstAfter = seriesAfter[0];
	const catBefore = firstBefore?.categories ?? firstBefore?.xValues;
	const catAfter = firstAfter?.categories ?? firstAfter?.xValues;
	const catBeforeTexts = catBefore?.cache ? cellTexts(catBefore.cache) : undefined;
	const catAfterTexts = catAfter?.cache ? cellTexts(catAfter.cache) : undefined;
	const hasCatLabels =
		catBeforeTexts &&
		catAfterTexts &&
		catBeforeTexts.size > 0 &&
		[...catBeforeTexts.values()].some((v) => v.trim() !== '');

	const commonSeriesCount = Math.min(seriesBefore.length, seriesAfter.length);
	const beforeFp: string[] = [];
	for (let i = 0; i < totalBefore; i++) {
		if (hasCatLabels) {
			beforeFp.push(catBeforeTexts.get(i) ?? '');
		} else {
			const parts: string[] = [];
			for (let s = 0; s < commonSeriesCount; s++) {
				const sb = seriesBefore[s];
				if (!sb) continue;
				const v = sb.values ?? sb.yValues;
				if (v?.cache) parts.push(cellTexts(v.cache).get(i) ?? '');
			}
			beforeFp.push(parts.join('\0'));
		}
	}

	const afterFp: string[] = [];
	for (let j = 0; j < totalAfter; j++) {
		if (hasCatLabels) {
			afterFp.push(catAfterTexts.get(j) ?? '');
		} else {
			const parts: string[] = [];
			for (let s = 0; s < commonSeriesCount; s++) {
				const sa = seriesAfter[s];
				if (!sa) continue;
				const v = sa.values ?? sa.yValues;
				if (v?.cache) parts.push(cellTexts(v.cache).get(j) ?? '');
			}
			afterFp.push(parts.join('\0'));
		}
	}

	const dp: number[][] = Array.from({ length: totalBefore + 1 }, () =>
		new Array<number>(totalAfter + 1).fill(0),
	);

	for (let i = 1; i <= totalBefore; i++) {
		for (let j = 1; j <= totalAfter; j++) {
			if (beforeFp[i - 1] === afterFp[j - 1]) {
				dp[i]![j] = dp[i - 1]![j - 1]! + 1;
			} else {
				dp[i]![j] = Math.max(dp[i - 1]![j]!, dp[i]![j - 1]!);
			}
		}
	}

	const survivingAfterToBefore = new Map<number, number>();
	const matchedBefore = new Set<number>();
	let bi = totalBefore;
	let aj = totalAfter;

	while (bi > 0 && aj > 0) {
		if (beforeFp[bi - 1] === afterFp[aj - 1]) {
			survivingAfterToBefore.set(aj - 1, bi - 1);
			matchedBefore.add(bi - 1);
			bi--;
			aj--;
		} else if (dp[bi - 1]![aj]! >= dp[bi]![aj - 1]!) {
			bi--;
		} else {
			aj--;
		}
	}

	if (survivingAfterToBefore.size < totalAfter) {
		let lastBefore = -1;
		for (let a = 0; a < totalAfter; a++) {
			if (!survivingAfterToBefore.has(a)) {
				for (let b = lastBefore + 1; b < totalBefore; b++) {
					if (!matchedBefore.has(b)) {
						survivingAfterToBefore.set(a, b);
						matchedBefore.add(b);
						lastBefore = b;
						break;
					}
				}
			} else {
				lastBefore = survivingAfterToBefore.get(a)!;
			}
		}
	}

	const removedRows: number[] = [];
	for (let b = 0; b < totalBefore; b++) {
		if (!matchedBefore.has(b)) removedRows.push(b);
	}

	return { removedRows, totalRowsBefore: totalBefore, survivingAfterToBefore };
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
	const alignment = before ? alignRows(space, before) : undefined;
	const seriesCount = allSeries(space).length;
	const beforeSeriesCount = before ? allSeries(before).length : seriesCount;
	const seriesRemoved = before !== undefined && seriesCount < beforeSeriesCount;

	const seen = new Set<string>();
	const writes: ChartWorkbookWrite[] = [];
	for (const series of allSeries(space)) {
		const isTx = (source: WritableSource) => source === series.tx?.reference;
		for (const source of writableSources(series)) {
			const { formula, cache } = source;
			const start = rangeStart(formula);
			if (seen.has(start)) continue;
			seen.add(start);
			const earlier = prior.get(start);
			const isDataRowRange = !isTx(source);
			const cells =
				alignment && isDataRowRange
					? chartShiftedCacheCellChanges(cache, earlier, alignment.survivingAfterToBefore)
					: chartCacheCellChanges(cache, earlier);
			// A range `before` did not read may hold anything (a series removed and brought back by
			// undo, data put back on a sheet it had left): the cache is all it holds.
			const only = before !== undefined && earlier === undefined;
			if (cells.length || !before || only || (alignment && isDataRowRange) || seriesRemoved)
				writes.push({
					formula,
					isNumeric: cache.type === 'number',
					cells,
					cellCount: chartCacheCellCount(cache),
					...(only ? { only } : {}),
					...(alignment && isDataRowRange
						? {
								removedRows: alignment.removedRows,
								totalRowsBefore: alignment.totalRowsBefore,
							}
						: {}),
					...(seriesRemoved ? { seriesCount } : {}),
				});
		}
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
