// Sliding-window rate limiter: at most `limit` hits per key within `windowMs`. `check` records the
// hit and returns `{ ok: true }`, or `{ ok: false, retryAfterSec }` when the key is over the limit
// (a refused hit is not recorded, so the wait is bounded by the window). Memory is bounded by
// pruning idle keys and capping the number of tracked keys.
export function createSlidingLimiter({ limit, windowMs, now = Date.now, maxKeys = 10000 }) {
	/** @type {Map<string, number[]>} */
	const hits = new Map();
	const prune = (t) => {
		for (const [key, times] of hits) {
			while (times.length > 0 && times[0] <= t - windowMs) times.shift();
			if (times.length === 0) hits.delete(key);
		}
	};
	return {
		check(key) {
			const t = now();
			let times = hits.get(key);
			if (times) while (times.length > 0 && times[0] <= t - windowMs) times.shift();
			if (!times) {
				if (hits.size >= maxKeys) prune(t);
				// Still full of live keys: fail closed for new ones rather than grow without bound.
				if (hits.size >= maxKeys) return { ok: false, retryAfterSec: Math.ceil(windowMs / 1000) };
				times = [];
				hits.set(key, times);
			}
			if (times.length >= limit) {
				const wait = Math.ceil((times[0] + windowMs - t) / 1000);
				return { ok: false, retryAfterSec: Math.max(1, wait) };
			}
			times.push(t);
			return { ok: true, retryAfterSec: 0 };
		},
		/** Whether `key` is at its limit now, without recording a hit (for counting only failures). */
		full(key) {
			const t = now();
			const times = hits.get(key);
			if (times) while (times.length > 0 && times[0] <= t - windowMs) times.shift();
			return (times?.length ?? 0) >= limit;
		},
		size: () => hits.size,
	};
}
