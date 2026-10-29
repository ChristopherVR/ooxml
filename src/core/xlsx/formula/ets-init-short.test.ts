// The seasonal start below 2 seasons (functions/ets-init-short.ts) against the starting-state maps
// recovered from Excel's impulse probes (ets-init14 to ets-init16): every probed series from the first
// seasonal length up to 2 seasons minus 1 point, seasons of 5 to 13.
import { describe, expect, it } from 'vitest';
import { chainStart } from './functions/ets-init-chain';
import { isShortSeasonLength, seasonStart } from './functions/ets-init-short';
import { recoveredMaps, type StartMap } from './ets-init-maps';

const maps = recoveredMaps().filter((g) => g.missing === 0 && isShortSeasonLength(g.n, g.m));
const unit = (n: number, i: number): number[] =>
	Array.from({ length: n }, (_v, j) => (i === j ? 1 : 0));

/** The map of `chainStart` on unit vectors: rows of z (level + season + trend) and of the trend. */
function ourMap(g: StartMap): { z: number[][]; trend: number[] } {
	const z: number[][] = Array.from({ length: g.m }, () => []);
	const trend: number[] = [];
	for (let i = 0; i < g.n; i++) {
		const init = chainStart(unit(g.n, i), g.m);
		if (!init) throw new Error(`no start for ${g.key}`);
		for (let k = 0; k < g.m; k++)
			(z[k] as number[]).push(init.level + (init.season[k] as number) + init.trend);
		trend.push(init.trend);
	}
	return { z, trend };
}

describe('seasonal start below 2 seasons', () => {
	it('covers the probed lengths', () => {
		const keys = new Set(maps.map((g) => g.key));
		for (const key of ['s5-9', 's6-10', 's7-12', 's8-14', 's9-17', 's10-18', 's12-22', 's13-25'])
			expect(keys.has(key)).toBe(true);
	});
	for (const g of maps) {
		it(`matches Excel for ${g.key} (${g.source})`, () => {
			const ours = ourMap(g);
			g.z.forEach((row, k) =>
				row.forEach((v, i) =>
					expect(Math.abs(v - ((ours.z[k] as number[])[i] as number))).toBeLessThan(1e-6),
				),
			);
			g.trend.forEach((v, i) => expect(Math.abs(v - (ours.trend[i] as number))).toBeLessThan(1e-8));
		});
	}

	// Seasons of 2 from 5 points use the 3-season rule on the series extended to 3 seasons (the first
	// length with a season for m = 2, see ets-init-threeplus.ts).
	for (const g of recoveredMaps().filter((q) => q.missing === 0 && q.m === 2 && q.n === 5)) {
		it(`matches Excel for ${g.key} (${g.source})`, () => {
			const ours = ourMap(g);
			g.z.forEach((row, k) =>
				row.forEach((v, i) =>
					expect(Math.abs(v - ((ours.z[k] as number[])[i] as number))).toBeLessThan(1e-6),
				),
			);
			g.trend.forEach((v, i) => expect(Math.abs(v - (ours.trend[i] as number))).toBeLessThan(1e-6));
		});
	}

	it('starts using the season where 0.7 n reaches m + 1', () => {
		const first: Record<number, number> = {
			2: 5,
			3: 6,
			4: 8,
			5: 9,
			6: 10,
			7: 12,
			8: 13,
			9: 15,
			10: 16,
			11: 18,
			12: 19,
			13: 20,
			14: 22,
			15: 23,
			16: 25,
		};
		for (const [m, n] of Object.entries(first)) expect(seasonStart(Number(m))).toBe(n);
	});
});
