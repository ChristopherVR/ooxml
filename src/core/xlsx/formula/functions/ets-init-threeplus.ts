// Excel's seasonal FORECAST.ETS start for 3 whole seasons plus more than `ceil(m / 2) + 1` points but
// fewer than 4 seasons (`3m + ceil(m / 2) + 2 <= n < 4m`, seasons of 4 and more), recovered from the
// impulse probes of ets-init11 and later (ets-init-threeplus.test.ts). It is the 3-season rule of
// `ets-init-three.ts` (`threeSeasonSeasonals`, applied unchanged to the longer series) with a shifted
// template taken off the states of the first positions. With `h = ceil(m / 2)`, `a = h + 1`
// and `D_k(y) = (y[m + k] - y[2m + k]) / 2 - (mean y[a + k .. a + k + m) - mean y[a + 2m + k .. a +
// 3m + k)) / 4`, the 3-season state `z[k]` (level plus season at position k, with position 0 the
// reference) becomes `z[k] - D_k` for the first `n - 3m - h - 1` positions (k = 0 up) and stays
// `z[k]` for the others; seasonal states are then taken relative to position 0 as always.
import { threeSeasonSeasonals } from './ets-init-three';

const mean = (xs: readonly number[]): number => xs.reduce((a, b) => a + b, 0) / (xs.length || 1);

/** Whether `threePlusSeasonals` is Excel's rule for `n` values with seasons of `m`. */
export function isThreePlusLength(n: number, m: number): boolean {
	return m >= 4 && n >= 3 * m + Math.ceil(m / 2) + 2 && n < 4 * m;
}

/** The shifted template that moves position k toward the 4-season chain. */
function shift(y: readonly number[], m: number, k: number): number {
	const a = Math.ceil(m / 2) + 1;
	const at = (i: number): number => y[i] as number;
	return (
		(at(m + k) - at(2 * m + k)) / 2 -
		(mean(y.slice(a + k, a + k + m)) - mean(y.slice(a + 2 * m + k, a + 3 * m + k))) / 4
	);
}

/** Seasonal states relative to position 0, for lengths where `isThreePlusLength` holds. */
export function threePlusSeasonals(y: readonly number[], m: number): number[] {
	const base = threeSeasonSeasonals(y, m);
	const moved = y.length - 3 * m - Math.ceil(m / 2) - 1;
	const z = base.map((s, k) => (k < moved ? s - shift(y, m, k) : s));
	const origin = z[0] as number;
	return z.map((s) => s - origin);
}
