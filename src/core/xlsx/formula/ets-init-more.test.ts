// The seasonal starts for 4 seasons plus a partial one (functions/ets-init-four.ts) and for 3 seasons
// plus more than ceil(m / 2) + 1 points (functions/ets-init-threeplus.ts) against the starting-state
// maps recovered from Excel's impulse probes, seasons of 4 to 10.
import { describe, expect, it } from 'vitest';
import { chainStart, isExactLength } from './functions/ets-init-chain';
import { isLowThreeLength, isThreePlusLength } from './functions/ets-init-threeplus';
import { isTwoPlusLength } from './functions/ets-init-three';
import { recoveredMaps } from './ets-init-maps';

const unit = (n: number, i: number): number[] =>
	Array.from({ length: n }, (_v, j) => (i === j ? 1 : 0));

const fourPlus = (n: number, m: number): boolean => n > 4 * m && n < 5 * m;
const maps = recoveredMaps().filter(
	(g) =>
		g.missing === 0 &&
		(fourPlus(g.n, g.m) ||
			isThreePlusLength(g.n, g.m) ||
			isTwoPlusLength(g.n, g.m) ||
			isLowThreeLength(g.n, g.m)),
);

describe('4 seasons plus a partial one, 3 seasons plus many points, 2 seasons plus a partial one', () => {
	it('covers the probed lengths', () => {
		const keys = new Set(maps.map((g) => g.key));
		for (const key of ['s4-17', 's5-22', 's6-27', 's7-32', 's8-35'])
			expect(keys.has(key)).toBe(true);
		for (const key of ['s7-27', 's8-30', 's8-31', 's9-35', 's10-39'])
			expect(keys.has(key)).toBe(true);
	});
	for (const g of maps) {
		it(`matches Excel for ${g.key} (${g.source})`, () => {
			for (let i = 0; i < g.n; i++) {
				const init = chainStart(unit(g.n, i), g.m);
				expect(init).toBeDefined();
				if (!init) return;
				g.z.forEach((row, k) =>
					expect(
						Math.abs((row[i] as number) - (init.level + (init.season[k] as number) + init.trend)),
					).toBeLessThan(1e-6),
				);
				expect(Math.abs((g.trend[i] as number) - init.trend)).toBeLessThan(1e-6);
			}
		});
	}

	it('is exact for a trend plus any seasonal pattern', () => {
		const base = [3, -1, 4, -2, 0.5, 2, -3, 1.5];
		for (const [m, n] of [
			[6, 26],
			[7, 30],
			[7, 27],
			[8, 30],
		] as const) {
			const pattern = base.slice(0, m);
			const y = Array.from({ length: n }, (_v, t) => 10 + 0.7 * t + (pattern[t % m] as number));
			const init = chainStart(y, m);
			expect(init?.trend).toBeCloseTo(0.7, 9);
			init?.season.forEach((s, k) =>
				expect(s - (init.season[0] as number)).toBeCloseTo(
					(pattern[k] as number) - (pattern[0] as number),
					9,
				),
			);
		}
	});

	it('marks the lengths exact', () => {
		expect(isExactLength(17, 4)).toBe(true);
		expect(isExactLength(23, 6)).toBe(true);
		expect(isExactLength(22, 6)).toBe(true);
		expect(isThreePlusLength(22, 6)).toBe(false);
		expect(isThreePlusLength(23, 6)).toBe(true);
		expect(isThreePlusLength(24, 6)).toBe(false);
		expect(isExactLength(26, 6)).toBe(true);
	});
});
