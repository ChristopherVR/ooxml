/**
 * Cache readers of the pptx chart adapter (`chart-from-neutral.ts`): they read the neutral chart
 * model's point caches (`ChartDataSource`) into the shapes `PptxChartData` keeps, with the rules
 * the object-tree readers they replace applied (`extractSeriesNumbersWithBlanks`,
 * `extractChartCategoryValues`, `extractMultiLevelCategoryValues`, `extractChartSeriesName`,
 * `parseChartDateCategories`): point text is trimmed, caches are expanded by `c:pt/@idx` to
 * `c:ptCount`, and a point without `@idx` is point 0.
 *
 * @module chart-from-neutral-data
 */
import type {
	ChartCachePoint,
	ChartDataCache,
	ChartDataSource,
	ChartPlotGroup,
	ChartText,
	ChartTitle,
} from '../../../chart/index';
import type { PptxChartDateCategories } from '../types';

const pointIndex = (point: ChartCachePoint) => point.index ?? 0;

/**
 * The cache's points by index (those `read` accepts) and its length: `declared` when it is longer
 * than the highest index, else the highest index plus one.
 */
function expandCache<T>(
	cache: ChartDataCache | undefined,
	read: (text: string) => T | undefined,
	declared: number | undefined,
): { byIndex: Map<number, T>; length: number } | undefined {
	if (!cache || cache.points.length === 0) return undefined;
	const byIndex = new Map<number, T>();
	let maxIndex = -1;
	for (const point of cache.points) {
		const index = pointIndex(point);
		const value = read(point.value.trim());
		if (value !== undefined) byIndex.set(index, value);
		maxIndex = Math.max(maxIndex, index);
	}
	const length = declared !== undefined && declared > maxIndex + 1 ? declared : maxIndex + 1;
	return { byIndex, length };
}

const readNumber = (text: string) => {
	const value = Number.parseFloat(text);
	return text.length > 0 && Number.isFinite(value) ? value : undefined;
};
const readText = (text: string) => (text.length > 0 ? text : undefined);
const isNumeric = (source: ChartDataSource | undefined) =>
	source?.kind === 'numRef' || source?.kind === 'numLit';

/** A numeric source expanded to its declared length; a missing or non-numeric point is null. */
export function neutralNumbersWithBlanks(source: ChartDataSource | undefined): (number | null)[] {
	if (!isNumeric(source)) return [];
	const expanded = expandCache(source?.cache, readNumber, source?.cache?.pointCount);
	if (!expanded || expanded.length <= 0) return [];
	return Array.from({ length: expanded.length }, (_, at) => expanded.byIndex.get(at) ?? null);
}

/**
 * A category list of a string (`numeric` false) or number (`numeric` true) source, blank slots kept
 * as empty strings; empty when the cache holds no text at all.
 */
export function neutralCategoryValues(
	source: ChartDataSource | undefined,
	numeric: boolean,
): string[] {
	if (!source?.kind || isNumeric(source) !== numeric || source.kind === 'multiLvlStrRef') return [];
	const expanded = expandCache(source.cache, readText, source.cache?.pointCount);
	if (!expanded || expanded.byIndex.size === 0) return [];
	return Array.from({ length: expanded.length }, (_, at) => expanded.byIndex.get(at) ?? '');
}

/**
 * A multi-level category source (`c:multiLvlStrRef`): the innermost level as the category list and
 * every level, the outer ones forward-filled across the leaves they span.
 */
export function neutralCategoryLevels(
	source: ChartDataSource | undefined,
): { categories: string[]; categoryLevels?: string[][] } | undefined {
	if (source?.kind !== 'multiLvlStrRef' || !source.levels?.length) return undefined;
	const levels = source.levels.map((level, levelIndex) => {
		const expanded = expandCache(level, readText, source.levelPointCount);
		const length = expanded?.length ?? Math.max(source.levelPointCount ?? 0, 0);
		const dense = Array.from({ length }, (_, at) => expanded?.byIndex.get(at) ?? '');
		if (levelIndex > 0) {
			let last = '';
			for (let at = 0; at < dense.length; at++) {
				if (dense[at]) last = dense[at];
				else dense[at] = last;
			}
		}
		return dense;
	});
	return { categories: levels[0] ?? [], ...(levels.length > 1 ? { categoryLevels: levels } : {}) };
}

/**
 * The chart-level category list: the first series of the first group that yields any supplies it
 * (multi-level, else string, else number categories, else the string form of its x values).
 */
export function neutralCategories(groups: ChartPlotGroup[]): {
	categories: string[];
	categoryLevels?: string[][];
} {
	let categoryLevels: string[][] | undefined;
	for (const group of groups) {
		const first = group.series[0];
		if (!first) continue;
		const multiLevel = neutralCategoryLevels(first.categories);
		if (multiLevel?.categoryLevels) categoryLevels = multiLevel.categoryLevels;
		const fromCategories = multiLevel
			? multiLevel.categories
			: neutralCategoryValues(first.categories, false);
		const fromNumbers =
			multiLevel || fromCategories.length ? [] : neutralCategoryValues(first.categories, true);
		const categories = fromCategories.length
			? fromCategories
			: fromNumbers.length
				? fromNumbers
				: neutralCategoryValues(first.xValues, false);
		if (categories.length > 0) return { categories, ...(categoryLevels ? { categoryLevels } : {}) };
	}
	return { categories: [], ...(categoryLevels ? { categoryLevels } : {}) };
}

/** The non-empty point texts of a cache, trimmed, in index order. */
function cachedTexts(cache: ChartDataCache | undefined): string[] {
	return (cache?.points ?? [])
		.slice()
		.sort((left, right) => pointIndex(left) - pointIndex(right))
		.map((point) => point.value.trim())
		.filter((value) => value.length > 0);
}

/** A series name (`c:tx`): the literal value, else the first cached text, else `Series`. */
export function neutralSeriesName(tx: ChartText | undefined): string {
	const literal = tx?.value?.trim();
	if (literal) return literal;
	return cachedTexts(tx?.reference?.cache)[0] ?? 'Series';
}

/**
 * The flat chart title: the first text run of the rich title, else the linked cell's cached text
 * (the full run list is `titleRuns`, read with the run styles).
 */
export function neutralTitleText(title: ChartTitle | undefined): string | undefined {
	const firstRun = title?.tx?.rich?.paragraphs.flatMap((paragraph) => paragraph.runs)[0];
	if (firstRun) return firstRun.text;
	return cachedTexts(title?.tx?.reference?.cache)[0] || undefined;
}

/** Numeric date categories of a date axis (`c:cat/c:numRef` of the first series). */
export function neutralDateCategories(
	source: ChartDataSource | undefined,
): PptxChartDateCategories | undefined {
	if (!isNumeric(source) || !source?.cache) return undefined;
	const cache = source.cache;
	const values = cache.points
		.slice()
		.sort((left, right) => pointIndex(left) - pointIndex(right))
		.map((point) => (point.value.trim() ? Number(point.value) : Number.NaN))
		.filter(Number.isFinite);
	if (values.length === 0) return undefined;
	const formatCode = cache.formatCode?.trim();
	return { values, ...(formatCode ? { formatCode } : {}) };
}
