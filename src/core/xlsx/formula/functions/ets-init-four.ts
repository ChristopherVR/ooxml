// Excel's seasonal FORECAST.ETS start for 4 whole seasons plus a partial one (`4m < n < 5m`, seasons of
// 4 and more), recovered from the impulse probes of ets-init4 to ets-init10 (ets-init-four.test.ts).
// It is the 4-season chain of the first `4m` values (`chainStates`) with a correction on the
// states of the times `a + q`, `a = floor(m / 2) - 1` and `q = 0 .. r - 1` (r = n - 4m points past
// the whole seasons): the state at time `a + q` gets `T(w)` for the window `w = y[q .. q + 4m)`,
//   T(w) = (B0 - B2) / 12 + (B1 - B3) / 6 - (w[a] + w[a + m] - 2 w[a + 2m]) / 6
// with `Bc` the mean of season `c` of the window. The correction is the same shifted template for
// every r and m checked (m 4 to 8, r 1 to 3 fully; see the test), and does not accumulate along the
// chain.
import { chainStates } from './ets-init-states';

const mean = (xs: readonly number[]): number => xs.reduce((a, b) => a + b, 0) / (xs.length || 1);

/** The correction template on a window of 4 whole seasons with the state at window index `a`. */
function template(w: readonly number[], m: number, a: number): number {
	const block = (c: number): number => mean(w.slice(c * m, (c + 1) * m));
	const at = (i: number): number => w[i] as number;
	return (
		(block(0) - block(2)) / 12 +
		(block(1) - block(3)) / 6 -
		(at(a) + at(a + m) - 2 * at(a + 2 * m)) / 6
	);
}

/** Seasonal states relative to position 0 for `4m < n < 5m`. */
export function fourSeasonPlusSeasonals(y: readonly number[], m: number): number[] {
	const r = y.length - 4 * m;
	const a = Math.floor(m / 2) - 1;
	const states = chainStates(y.slice(0, 4 * m), m, 4);
	for (let q = 0; q < r; q++) {
		const pos = (a + q) % m;
		states[pos] = (states[pos] as number) + template(y.slice(q, q + 4 * m), m, a);
	}
	const origin = states[0] as number;
	return states.map((s) => s - origin);
}
