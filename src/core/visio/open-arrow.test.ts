import { describe, expect, it } from 'vitest';
import { visioOpenArrowExtent, visioOpenArrowPath } from './open-arrow';
import evidence from './__fixtures__/open-arrows-native.json';

describe('native open-arrow geometry at unit drawing scale', () => {
	it.each(evidence.cases)('matches size $size at line width $lineWidth', (row) => {
		expect(visioOpenArrowExtent(row.size, row.lineWidth)).toBeCloseTo(row.extent, 12);
	});
	it.each(evidence.cases)('matches path code $code size $size weight $lineWidth', (row) => {
		const path = visioOpenArrowPath(row.code, row.size, row.lineWidth)!;
		expect(path.match(/[A-Z]/g)).toEqual(row.nativeGlyph.match(/[A-Z]/g));
		const values = (d: string) => d.match(/-?\d+(?:\.\d+)?/g)!.map(Number);
		const actual = values(path);
		values(row.nativeGlyph).forEach((n, i) => {
			expect(actual[i]).toBeCloseTo(n * row.extent * (i % 2 ? 1 : -1), 12);
		});
	});
	it('leaves filled and unsupported glyphs to other render paths', () => {
		for (const code of [0, 2, 4, 5, 6, 8, 254])
			expect(visioOpenArrowPath(code, 2, 0.01)).toBeUndefined();
	});
	it('bounds malformed sizes and weights without emitting non-finite coordinates', () => {
		expect(visioOpenArrowExtent(NaN, Infinity)).toBe(0.035);
		expect(visioOpenArrowExtent(-1, -1)).toBe(0.02);
		expect(visioOpenArrowExtent(99, 0.01)).toBe(0.26);
	});
});
