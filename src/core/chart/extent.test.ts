import { describe, expect, it } from 'vitest';
import { extentOf } from './extent';

describe('extentOf', () => {
	it('matches Math.min and Math.max over the same values', () => {
		const cases = [
			[],
			[3],
			[2, -1, 5],
			[0, -0],
			[-0, 0],
			[1, Number.NaN, 2],
			[Infinity, -Infinity],
		];
		for (const values of cases)
			expect(extentOf(values)).toEqual({ min: Math.min(...values), max: Math.max(...values) });
	});

	it('takes more values than a call can pass as arguments', () => {
		const values = Array.from({ length: 200_000 }, (_, i) => (i * 7919) % 200_000);
		values[123_456] = -3;
		values[45_678] = 250_000;
		expect(extentOf(values)).toEqual({ min: -3, max: 250_000 });
	});
});
