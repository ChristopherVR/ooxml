// What the recovered starting-state maps of Excel's FORECAST.ETS (ets-init-maps.ts) establish,
// beyond the filter itself:
//  1. Without a season the starting trend is exactly `shortSeriesTrend`: the plain least-squares
//     slope from 7 values (checked to 16), and for 4 to 6 values the slope of the line through the
//     data extended to 7 points, 3 values being the plain slope too.
//  2. With a season the starting trend ignores any periodic pattern added to the data (the weights
//     of every position sum to 0) and the state of position k moves by exactly that pattern.
//  3. Below 3 seasons Excel runs its 3-season algorithm on the series extended by copies of the
//     first points, each moved up by `n` times the slope of a trend-plus-season-dummies fit
//     (checked for season 3 with 6 values and season 4 with 8; 7 points is the minimum for season
//     2, which has 6 values for 3 seasons). Its 3-season map itself is the part not yet explained.
//  4. From 4 seasons, for seasons of 2 and 3, the state of position k follows the one before it:
//     `z[k] = z[k-1] + P[k] - P[k-1] - b[k-1]`, P being the mean of the first `C - h` seasons of
//     that position and b the half-split slope of position `k - 1` (mean of its last `h` seasons
//     minus mean of its first `h`, over the distance between those means, `h = floor(C / 2)`);
//     `z[0]` is the first value. Larger seasons use the same chain with the slope of position k - floor(m / 2) (ets-init-chain.ts).
//     The starting trend is the least-squares slope of the series without those states (ets-init-chain.test.ts).
import { describe, expect, it } from 'vitest';
import { shortSeriesTrend, solveSystem } from './functions/ets-init';
import { recoveredMaps, type StartMap } from './ets-init-maps';

const maps = recoveredMaps().filter((g) => g.missing === 0);
const pick = (key: string): StartMap => {
	const found = maps.find((g) => g.key === key);
	if (!found) throw new Error(`no map ${key}`);
	return found;
};
const unit = (n: number, i: number): number[] =>
	Array.from({ length: n }, (_v, j) => (i === j ? 1 : 0));
const close = (a: readonly number[], b: readonly number[], tol = 1e-8): void =>
	a.forEach((v, i) => expect(Math.abs(v - (b[i] as number))).toBeLessThan(tol));
/** Applies a map to data given as weight vectors over a smaller series. */
const compose = (rows: number[][], data: number[][]): number[][] =>
	rows.map((w) =>
		(data[0] as number[]).map((_v, j) =>
			data.reduce((s, d, i) => s + (w[i] as number) * (d[j] as number), 0),
		),
	);

/** Weights of the slope of `y = a + b t + s[t mod m]` fitted by least squares to `n` values. */
function dummySlope(n: number, m: number): number[] {
	const rows = Array.from({ length: n }, (_v, t) => [
		1,
		t,
		...Array.from({ length: m - 1 }, (_x, j) => (t % m === j + 1 ? 1 : 0)),
	]);
	const k = m + 1;
	const gram = (): number[][] =>
		Array.from({ length: k }, (_v, i) =>
			Array.from({ length: k }, (_w, j) =>
				rows.reduce((s, r) => s + (r[i] as number) * (r[j] as number), 0),
			),
		);
	return Array.from({ length: n }, (_v, i) => {
		const rhs = Array.from({ length: k }, (_w, j) => (rows[i] as number[])[j] as number);
		return solveSystem(gram(), rhs)[1] as number;
	});
}

describe('Excel FORECAST.ETS starting-state maps', () => {
	it('uses the extended-to-7 least-squares slope without a season, 3 to 16 values', () => {
		const plain = maps.filter((g) => g.m === 0);
		expect(plain.map((g) => g.n).sort((a, b) => a - b)).toEqual([
			3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14, 16,
		]);
		for (const g of plain) {
			const ours = Array.from({ length: g.n }, (_v, i) => shortSeriesTrend(unit(g.n, i)));
			close(g.trend, ours);
			close(g.z[0] as number[], unit(g.n, 0));
		}
	});

	it('keeps a periodic pattern out of the trend', () => {
		// s2-4 of the first file is a series where Excel finds no season (gamma stays at its floor).
		const seasonal = maps.filter((g) => g.m >= 2 && g.n >= 2 * g.m && g.key !== 's2-4');
		expect(seasonal.length).toBeGreaterThan(20);
		for (const g of seasonal)
			for (let k = 0; k < g.m; k++) {
				let trend = 0;
				let own = 0;
				for (let t = k; t < g.n; t += g.m) {
					trend += g.trend[t] as number;
					own += (g.z[k] as number[])[t] as number;
				}
				// 1e-4: the probes with an impulse of 1 lose a few digits in the recovery.
				expect(Math.abs(trend)).toBeLessThan(1e-4);
				expect(Math.abs(own - 1)).toBeLessThan(1e-4);
			}
	});

	it('extends a short seasonal series to 3 seasons with copies moved by n slopes', () => {
		for (const [m, n, key] of [
			[3, 6, 's3-9'],
			[4, 8, 's4-12'],
			[2, 6, 's2-7'],
		] as const) {
			const small = pick(`s${m}-${n}`);
			const big = pick(key);
			const slope = dummySlope(n, m);
			const padded = [
				...Array.from({ length: n }, (_v, i) => unit(n, i)),
				...Array.from({ length: big.n - n }, (_v, j) =>
					unit(n, j).map((v, i) => v + n * (slope[i] as number)),
				),
			];
			const rows = [...big.z, big.trend];
			const expected = compose(rows, padded);
			[...small.z, small.trend].forEach((row, r) => close(row, expected[r] as number[]));
		}
	});

	it('chains the position states from 4 seasons for seasons of 2 and 3', () => {
		const cases = [
			['s2-8', 2, 8],
			['s2-10', 2, 10],
			['s3-12', 3, 12],
			['s3-15', 3, 15],
		] as const;
		for (const [key, m, n] of cases) {
			const g = pick(key);
			const C = n / m;
			const h = Math.floor(C / 2);
			const H = C - h;
			const gap = (C - h) * m;
			const at = (k: number, c: number): number[] => unit(n, k + c * m);
			const mean = (k: number, from: number, to: number): number[] =>
				Array.from({ length: to - from }, (_v, i) => at(k, from + i)).reduce(
					(s, v) => s.map((x, j) => x + (v[j] as number) / (to - from)),
					new Array<number>(n).fill(0),
				);
			const sub = (a: number[], b: number[]): number[] => a.map((v, i) => v - (b[i] as number));
			const slopeOf = (k: number): number[] =>
				sub(mean(k, C - h, C), mean(k, 0, h)).map((v) => v / gap);
			let z = unit(n, 0);
			close(g.z[0] as number[], z);
			for (let k = 1; k < m; k++) {
				const step = sub(sub(mean(k, 0, H), mean(k - 1, 0, H)), slopeOf(k - 1));
				z = z.map((v, i) => v + (step[i] as number));
				close(g.z[k] as number[], z);
			}
		}
	});

	it('looks at no more than the first 5 seasons of 2 when setting position states', () => {
		const ten = pick('s2-10');
		for (const key of ['s2-12', 's2-14']) {
			const g = pick(key);
			ten.z.forEach((row, k) => {
				const longer = g.z[k] as number[];
				close(longer.slice(0, 10), row);
				close(longer.slice(10), new Array(g.n - 10).fill(0));
			});
		}
	});
});
