import { describe, expect, it } from 'vitest';

import { isFiniteNumber } from './finite';

describe('isFiniteNumber', () => {
	it('accepts finite numbers only, as Number.isFinite does', () => {
		for (const value of [0, -1.5, 1e300]) expect(isFiniteNumber(value)).toBe(true);
		for (const value of [NaN, Infinity, -Infinity, undefined, null, '1', {}]) {
			expect(isFiniteNumber(value)).toBe(Number.isFinite(value));
			expect(isFiniteNumber(value)).toBe(false);
		}
	});
});
