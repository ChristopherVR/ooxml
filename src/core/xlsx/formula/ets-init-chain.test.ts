// The exact seasonal start (functions/ets-init-chain.ts) against the starting-state maps recovered
// from Excel's impulse probes: every probed series of 4 or 5 whole seasons (seasons 2 to 6), and the
// seasons of 2 beyond 5 seasons, where Excel looks at the first 5 only.
import { describe, expect, it } from 'vitest';
import { chainStart, isExactLength } from './functions/ets-init-chain';
import { recoveredMaps, type StartMap } from './ets-init-maps';

const maps = recoveredMaps().filter((g) => g.missing === 0);
const unit = (n: number, i: number): number[] =>
	Array.from({ length: n }, (_v, j) => (i === j ? 1 : 0));

/** The map of `chainStart` on unit vectors: rows of z (level + season + trend) and of the trend. */
function ourMap(g: StartMap): { z: number[][]; trend: number[] } | undefined {
	const z: number[][] = Array.from({ length: g.m }, () => []);
	const trend: number[] = [];
	for (let i = 0; i < g.n; i++) {
		const init = chainStart(unit(g.n, i), g.m);
		if (!init) return undefined;
		for (let k = 0; k < g.m; k++)
			(z[k] as number[]).push(init.level + (init.season[k] as number) + init.trend);
		trend.push(init.trend);
	}
	return { z, trend };
}

describe('exact seasonal start', () => {
	const covered = maps.filter((g) => g.m >= 2 && isExactLength(g.n, g.m) && g.n >= 4 * g.m);
	it('covers the probed whole-season series', () => {
		expect(covered.length).toBeGreaterThanOrEqual(12);
	});
	for (const g of covered) {
		it(`matches Excel for ${g.key} (${g.source})`, () => {
			const ours = ourMap(g);
			expect(ours).toBeDefined();
			if (!ours) return;
			g.z.forEach((row, k) =>
				row.forEach((v, i) =>
					expect(Math.abs(v - ((ours.z[k] as number[])[i] as number))).toBeLessThan(1e-5),
				),
			);
			g.trend.forEach((v, i) => expect(Math.abs(v - (ours.trend[i] as number))).toBeLessThan(1e-5));
		});
	}
	it('leaves other lengths to the fallback', () => {
		expect(chainStart([1, 2, 3, 4, 5, 6, 7, 8, 9], 3)).toBeUndefined();
		expect(
			chainStart(
				Array.from({ length: 7 }, (_v, i) => i),
				4,
			),
		).toBeUndefined();
	});
});
