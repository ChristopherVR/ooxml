// Reading cached chart values (`c:numCache`, `c:strCache`, literals) as dense arrays, the way the
// producing application laid them out: missing points are empty strings, numeric caches yield
// numbers where the cached text is a finite number.
import type { ChartDataCache, ChartDataSource } from './model-series';
import { chartDensePoints, chartPointsLength } from './point-limit';

/** The points of a cache by index, holes filled with `''`, numbers parsed in numeric caches. */
export function chartCacheValues(cache: ChartDataCache | undefined): (string | number)[] {
	if (!cache) return [];
	const numeric = cache.type === 'number';
	return chartDensePoints<ChartDataCache['points'][number], string | number>(
		cache,
		({ value }) =>
			numeric && value.trim() !== '' && Number.isFinite(Number(value)) ? Number(value) : value,
		'',
	);
}

/** How many slots {@link chartCacheValues} returns, without building them. */
export const chartCachePointCount = (cache: ChartDataCache | undefined): number =>
	cache ? chartPointsLength(cache) : 0;

/** The cached values of a data source: its cache or literal, else the innermost category level. */
export const chartSourceValues = (source: ChartDataSource | undefined): (string | number)[] =>
	chartCacheValues(source?.cache ?? source?.levels?.[0]);
