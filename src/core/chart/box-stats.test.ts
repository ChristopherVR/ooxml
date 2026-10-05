import { describe, expect, it } from 'vitest';
import { computeBoxStats, groupRowsByCategory } from './box-stats.js';

describe('category observation grouping', () => {
	it('keeps first appearance order and groups noncontiguous and empty category labels', () => {
		const result = groupRowsByCategory(['B', '', 'A', 'B', '', 'A']);
		expect(result.uniqueCategories).toEqual(['B', '', 'A']);
		expect([...result.rowIndexesByCategory]).toEqual([
			['B', [0, 3]],
			['', [1, 4]],
			['A', [2, 5]],
		]);
	});
});

describe('computeBoxStats', () => {
	it('returns undefined for fewer than two values', () => {
		expect(computeBoxStats([5])).toBeUndefined();
		expect(computeBoxStats([])).toBeUndefined();
	});

	it('computes the five-number summary with floor-index quartiles', () => {
		// sorted: [10, 20, 30, 40] (n=4): q1 idx floor(1)=1 ->20, med idx floor(2)=2 ->30, q3 idx floor(3)=3 ->40.
		const stats = computeBoxStats([40, 10, 30, 20]);
		expect(stats).toStrictEqual({ min: 10, q1: 20, median: 30, q3: 40, max: 40 });
	});

	it('sorts input before computing quartiles', () => {
		const stats = computeBoxStats([100, 1, 50, 25, 75]);
		expect(stats?.min).toBe(1);
		expect(stats?.max).toBe(100);
	});

	it('orders min <= q1 <= median <= q3 <= max', () => {
		const stats = computeBoxStats([3, 1, 4, 1, 5, 9, 2, 6]);
		expect(stats).toBeDefined();
		if (stats) {
			expect(stats.min).toBeLessThanOrEqual(stats.q1);
			expect(stats.q1).toBeLessThanOrEqual(stats.median);
			expect(stats.median).toBeLessThanOrEqual(stats.q3);
			expect(stats.q3).toBeLessThanOrEqual(stats.max);
		}
	});

	it('distinguishes inclusive and exclusive interpolated quartiles', () => {
		const values = [1, 2, 3, 4, 5, 6, 7, 8];
		expect(computeBoxStats(values, 'inclusive')).toMatchObject({ q1: 2.75, q3: 6.25 });
		expect(computeBoxStats(values, 'exclusive')).toMatchObject({ q1: 2.25, q3: 6.75 });
	});
});
