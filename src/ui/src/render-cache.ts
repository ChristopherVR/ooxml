export const RENDER_CACHE_MAX_BYTES = 8 * 1024 * 1024;
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
	if (bytesOf(key, value) > RENDER_CACHE_MAX_BYTES) return;
	cache.set(key, value);
	let bytes = 0;
	for (const [entryKey, entryValue] of cache) bytes += bytesOf(entryKey, entryValue);
	while (cache.size > maxEntries || bytes > RENDER_CACHE_MAX_BYTES) {
		const first = cache.entries().next().value;
		if (!first) break;
		bytes -= bytesOf(first[0], first[1]);
		cache.delete(first[0]);
	}
}
