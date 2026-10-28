/** Conservative UTF-16 budget, including source URL keys and generated data URLs. */

let epoch = 0;

/** Capture before async image work; compare again before writing its result. */
export function getPptxRenderCacheEpoch(): number {
	return epoch;
}

export function invalidatePptxRenderCacheEpoch(): void {
	epoch++;
}

export { RENDER_CACHE_MAX_BYTES, setBoundedRenderCache } from '../../render-cache';
