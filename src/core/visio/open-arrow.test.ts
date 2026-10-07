import { describe, expect, it } from 'vitest';
import { visioOpenArrowExtent } from './open-arrow.js';
import evidence from './__fixtures__/open-arrows-native.json';

describe('native code-9 arrow dimensions at unit drawing scale', () => {
	it.each(evidence.cases)('matches size $size at line width $lineWidth', (row) => {
		expect(visioOpenArrowExtent(row.size, row.lineWidth)).toBeCloseTo(row.extent, 12);
	});
	it('bounds malformed sizes and weights without emitting non-finite coordinates', () => {
		expect(visioOpenArrowExtent(NaN, Infinity)).toBe(0.035);
		expect(visioOpenArrowExtent(-1, -1)).toBe(0.02);
		expect(visioOpenArrowExtent(99, 0.01)).toBe(0.26);
	});
});
