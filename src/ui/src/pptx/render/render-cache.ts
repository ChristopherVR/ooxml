/** Conservative UTF-16 budget, including source URL keys and generated data URLs. */
const MAX_BYTES = 8 * 1024 * 1024;
let epoch = 0;

/** Capture before async image work; compare again before writing its result. */
export function getPptxRenderCacheEpoch(): number {
	return epoch;
}

export function invalidatePptxRenderCacheEpoch(): void {
	epoch++;
}

/** FIFO retention limit; oversized results still render, but are not memoised. */
export function setBoundedRenderCache<T>(
	cache: Map<string, T>,
	key: string,
	value: T,
	maxEntries: number,
): void {
	const bytesOf = (entryKey: string, entryValue: T): number =>
		(entryKey.length + (typeof entryValue === 'string' ? entryValue.length : 0)) * 2;
	cache.delete(key);
	if (bytesOf(key, value) > MAX_BYTES) return;
	cache.set(key, value);
	let bytes = 0;
	for (const [entryKey, entryValue] of cache) bytes += bytesOf(entryKey, entryValue);
	while (cache.size > maxEntries || bytes > MAX_BYTES) {
		const first = cache.entries().next().value;
		if (!first) break;
		bytes -= bytesOf(first[0], first[1]);
		cache.delete(first[0]);
	}
}
