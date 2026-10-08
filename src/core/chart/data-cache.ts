// Reading cached chart values (`c:numCache`, `c:strCache`, literals) as dense arrays, the way the
// producing application laid them out: missing points are empty strings, numeric caches yield
// numbers where the cached text is a finite number.
import type { ChartDataCache, ChartDataSource } from './model-series';

/** The points of a cache by index, holes filled with `''`, numbers parsed in numeric caches. */
export function chartCacheValues(cache: ChartDataCache | undefined): (string | number)[] {
	if (!cache) return [];
	const numeric = cache.type === 'number';
	const out: (string | number)[] = [];
	for (const point of cache.points) {
		const text = point.value;
		out[point.index ?? out.length] =
			numeric && text.trim() !== '' && Number.isFinite(Number(text)) ? Number(text) : text;
	}
	const size = Math.max(cache.pointCount ?? 0, out.length);
	for (let index = 0; index < size; index++) if (out[index] === undefined) out[index] = '';
	return out;
}

/** The cached values of a data source: its cache or literal, else the innermost category level. */
export const chartSourceValues = (source: ChartDataSource | undefined): (string | number)[] =>
	chartCacheValues(source?.cache ?? source?.levels?.[0]);
