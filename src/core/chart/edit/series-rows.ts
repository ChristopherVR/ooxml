// Edits of single points and category rows, on a chart's sparse caches: every point the file holds
// stays, wherever a view of it ends, and a declared range keeps its count. Semantics follow
// PowerPoint's chart operations (`updateChartDataPoint`, `addChartCategory`, `removeChartCategory`).
import type { ChartSeriesValue } from '../data-sheet';
import type { ChartSpace } from '../model';
import type { ChartCachePoint, ChartDataCache, ChartDataSource } from '../model-series';
import {
	chartCacheWithPoint,
	chartCacheWithRow,
	chartCacheWithoutRow,
	chartPointsLength,
} from '../point-limit';
import { chartSeriesEntries, chartSeriesEntryAt, editChartSpace } from './draft';
import { chartCategoryKey, chartValueKey } from './series-keys';
import type { ChartCategoryValue } from './series-data';

/** How many rows a source reaches: its own cache or longest category level. */
const sourceRows = (source: ChartDataSource | undefined): number =>
	Math.max(
		0,
		...[...(source?.cache ? [source.cache] : []), ...(source?.levels ?? [])].map((cache) =>
			chartPointsLength(cache, source?.levels ? source.levelPointCount : cache.pointCount),
		),
	);

/** The rows of a chart's data: the most any series' categories or values reach. */
export function chartRowCount(space: ChartSpace): number {
	return Math.max(
		0,
		...chartSeriesEntries(space).flatMap(({ group, series }) => [
			sourceRows(series[chartCategoryKey(group)]),
			sourceRows(series[chartValueKey(group)]),
		]),
	);
}

const asPoint = (slot: number, value: string | number | null): ChartCachePoint | undefined =>
	value === null || value === '' ? undefined : { index: slot, value: String(value) };

/** `source` as a number source: the same reference and points, with a numeric cache. */
function numericSource(source: ChartDataSource | undefined): ChartDataSource {
	const formula = source?.kind === 'numRef' ? source.formula : undefined;
	return {
		kind: formula !== undefined ? 'numRef' : 'numLit',
		...(formula !== undefined ? { formula } : {}),
		cache: {
			points: [],
			...source?.cache,
			type: 'number',
			formatCode: source?.cache?.formatCode ?? 'General',
		},
	};
}

const emptyTextCache = (): ChartDataCache => ({ type: 'string', points: [] });

/** `source` with `edit` applied to its cache, or to every level of a category hierarchy. */
function withCaches(
	source: ChartDataSource,
	edit: (cache: ChartDataCache) => ChartDataCache,
): ChartDataSource {
	if (source.levels) return { ...source, levels: source.levels.map(edit) };
	return source.cache ? { ...source, cache: edit(source.cache) } : source;
}

/** A category source with a row at `slot` labelled `name`; text turns a numeric cache to text. */
function withCategoryRow(
	source: ChartDataSource,
	slot: number,
	name: ChartCategoryValue,
): ChartDataSource {
	const point = asPoint(slot, name);
	if (source.levels)
		return {
			...withCaches(source, (level) =>
				chartCacheWithRow(level, slot, level === source.levels![0] ? point : undefined),
			),
			...(source.levelPointCount === undefined
				? {}
				: { levelPointCount: Math.max(source.levelPointCount, slot + 1) }),
		};
	if (!source.cache) return { ...source, cache: chartCacheWithRow(emptyTextCache(), slot, point) };
	if (source.cache.type !== 'number' || typeof name !== 'string')
		return withCaches(source, (cache) => chartCacheWithRow(cache, slot, point));
	const { formatCode: _numeric, ...cache } = source.cache;
	const text: ChartDataCache = {
		...cache,
		type: 'string',
		points: cache.points.map(({ formatCode: _format, ...kept }) => kept),
	};
	return {
		...source,
		kind: source.formula !== undefined ? 'strRef' : 'strLit',
		cache: chartCacheWithRow(text, slot, point),
	};
}

/** A source without the row at `slot`; a complete hierarchy declares one row fewer. */
function withoutRow(source: ChartDataSource, slot: number): ChartDataSource {
	const next = withCaches(source, (cache) => chartCacheWithoutRow(cache, slot));
	if (
		next.levelPointCount !== undefined &&
		slot < next.levelPointCount &&
		!source.levels?.some((level) => level.viewLength !== undefined)
	)
		next.levelPointCount -= 1;
	return next;
}

/**
 * Sets the value in row `point` of a series, whatever its cache reaches: the row is that of the
 * chart's data, which can lie past the end of a shorter series. Every other point stays as written.
 */
export function setChartSeriesPoint(
	space: ChartSpace,
	index: number,
	point: number,
	value: ChartSeriesValue,
): ChartSpace {
	if (!Number.isInteger(point) || point < 0)
		throw new RangeError(`No chart point at index ${String(point)}`);
	if (value !== null && !Number.isFinite(value))
		throw new RangeError('Chart values must be finite numbers or null');
	return editChartSpace(space, (draft) => {
		const { group, series } = chartSeriesEntryAt(draft, index);
		const key = chartValueKey(group);
		const source = numericSource(series[key]);
		const cache = source.cache!;
		const previous = cache.points.find((item) => item.index === point);
		const written = asPoint(point, value);
		series[key] = {
			...source,
			cache: chartCacheWithPoint(cache, point, written && { ...previous, ...written }),
		};
	});
}

/** Sets one point of a series; the point must lie within the series. */
export function setChartDataPointValue(
	space: ChartSpace,
	index: number,
	point: number,
	value: ChartSeriesValue,
): ChartSpace {
	const { group, series } = chartSeriesEntryAt(space, index);
	if (point >= sourceRows(series[chartValueKey(group)]))
		throw new RangeError(`No chart point at index ${String(point)}`);
	return setChartSeriesPoint(space, index, point, value);
}

/**
 * Appends a category to every series, in the row after all the chart's data (or in `row`), with
 * `value` (PowerPoint: 0) as each series' point. A series that already has a point in that row
 * keeps it, and a declared range that reaches the row keeps its count. The row's x value of a
 * scatter or bubble series is left blank, as a new worksheet row is, and so is a bubble size.
 */
export function addChartCategory(
	space: ChartSpace,
	name: ChartCategoryValue,
	value: ChartSeriesValue = 0,
	row: number = chartRowCount(space),
): ChartSpace {
	return editChartSpace(space, (draft) => {
		for (const { group, series } of chartSeriesEntries(draft)) {
			const key = chartCategoryKey(group);
			const names = series[key];
			// A label is no x value: the row's x stays blank rather than turning the x values to text.
			if (names) series[key] = withCategoryRow(names, row, key === 'xValues' ? null : name);
			if (series.bubbleSizes)
				series.bubbleSizes = withCaches(series.bubbleSizes, (cache) =>
					chartCacheWithRow(cache, row, undefined),
				);
			const source = numericSource(series[chartValueKey(group)]);
			series[chartValueKey(group)] = {
				...source,
				cache: chartCacheWithRow(source.cache!, row, asPoint(row, value)),
			};
		}
	});
}

/** Removes a category row from every series, with its points, bubble sizes and point formatting. */
export function removeChartCategory(space: ChartSpace, category: number): ChartSpace {
	if (!Number.isInteger(category) || category < 0 || category >= chartRowCount(space))
		throw new RangeError(`No chart category at index ${String(category)}`);
	return editChartSpace(space, (draft) => {
		for (const { group, series } of chartSeriesEntries(draft)) {
			for (const key of [chartCategoryKey(group), chartValueKey(group), 'bubbleSizes'] as const) {
				const source = series[key];
				if (source) series[key] = withoutRow(source, category);
			}
			series.dataPoints = series.dataPoints
				.filter((p) => p.index !== category)
				.map((p) =>
					p.index !== undefined && p.index > category ? { ...p, index: p.index - 1 } : p,
				);
		}
	});
}
