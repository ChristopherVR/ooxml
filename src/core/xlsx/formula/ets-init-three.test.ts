// The 3-season (and exactly 2-season) seasonal start (functions/ets-init-three.ts) against the
// starting-state maps recovered from Excel's impulse probes: every probed series of 3 seasons plus up
// to m - 2 points and of 2 seasons, for seasons of 4 to 12.
import { describe, expect, it } from 'vitest';
import {
	extendTwoSeasons,
	isThreeSeasonLength,
	isTwoSeasonLength,
	threeSeasonSeasonals,
} from './functions/ets-init-three';
import { chainStart } from './functions/ets-init-chain';
import { recoveredMaps, type StartMap } from './ets-init-maps';

const maps = recoveredMaps().filter((g) => g.missing === 0 && g.m >= 4);
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

describe('3-season seasonal start', () => {
	const covered = maps.filter((g) => isThreeSeasonLength(g.n, g.m) || isTwoSeasonLength(g.n, g.m));
	it('covers the probed lengths', () => {
		const keys = new Set(covered.map((g) => g.key));
		for (const key of ['s4-12', 's4-13', 's4-14', 's5-15', 's5-16', 's5-17', 's7-21', 's12-36']) {
			expect(keys.has(key)).toBe(true);
		}
		for (const key of ['s4-8', 's5-10', 's6-12']) expect(keys.has(key)).toBe(true);
	});
	for (const g of covered) {
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

	it('is exact for a trend plus any seasonal pattern', () => {
		const m = 5;
		const n = 17;
		const pattern = [3, -1, 4, -2, 0.5];
		const y = Array.from({ length: n }, (_v, t) => 10 + 0.7 * t + (pattern[t % m] as number));
		const rel = threeSeasonSeasonals(y, m);
		rel.forEach((s, k) =>
			expect(s).toBeCloseTo((pattern[k] as number) - (pattern[0] as number), 9),
		);
		const init = chainStart(y, m);
		expect(init?.trend).toBeCloseTo(0.7, 9);
	});

	it('extends two seasons by the dummy-regression slope', () => {
		const y = [1, 2, 3, 4, 2, 3, 4, 5];
		const ext = extendTwoSeasons(y, 4);
		expect(ext).toHaveLength(12);
		// y[t] = pattern + t / 4: the extension continues the line.
		expect(ext[8]).toBeCloseTo(3, 9);
		expect(ext[11]).toBeCloseTo(6, 9);
	});

	it('leaves 4 seasons less one point and seasons below 4 to the fallback', () => {
		expect(isThreeSeasonLength(15, 4)).toBe(false);
		expect(isThreeSeasonLength(9, 3)).toBe(false);
		expect(isTwoSeasonLength(6, 3)).toBe(false);
	});
});
