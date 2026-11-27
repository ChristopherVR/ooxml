/** Seeded randomness for the generators: the same seed always gives the same sequence. */

export interface Random {
	/** A float in [0, 1). */
	next(): number;
	/** An integer in [min, max]. */
	int(min: number, max: number): number;
	/** A float in [min, max). */
	range(min: number, max: number): number;
	pick<T>(items: readonly T[]): T;
	/** A normally distributed float (Box-Muller). */
	normal(mean: number, deviation: number): number;
}

/** mulberry32: small, fast and identical on every engine. */
export function seeded(seed: number): Random {
	let state = seed >>> 0;
	const next = () => {
		state = (state + 0x6d2b79f5) >>> 0;
		let t = state;
		t = Math.imul(t ^ (t >>> 15), t | 1);
		t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
		return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
	};
	const range = (min: number, max: number) => min + next() * (max - min);
	return {
		next,
		range,
		int: (min, max) => Math.floor(range(min, max + 1)),
		pick: (items) => items[Math.floor(next() * items.length)]!,
		normal: (mean, deviation) => {
			const u = Math.max(next(), 1e-12);
			const v = next();
			return mean + deviation * Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v);
		},
	};
}

/** Rounds to `digits` decimals, so stored numbers do not carry float noise. */
export function round(value: number, digits = 2): number {
	const factor = 10 ** digits;
	return Math.round(value * factor) / factor;
}
