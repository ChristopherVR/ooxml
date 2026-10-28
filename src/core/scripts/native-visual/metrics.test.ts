import { describe, expect, it } from 'vitest';
import { comparePixels } from './metrics';

describe('native visual comparison evidence', () => {
	it('composites alpha without treating transparent black as foreground', () => {
		const a = new Uint8ClampedArray([0, 0, 0, 0, 0, 0, 0, 255]);
		const b = new Uint8ClampedArray([255, 255, 255, 255, 0, 0, 0, 255]);
		expect(comparePixels(a, b)).toMatchObject({
			normalizedMeanAbsoluteError: 0,
			changedPixelRatio: 0,
			foregroundIntersectionOverUnion: 1,
			nativeForegroundRatio: 0.5,
		});
	});
	it('does not pass blank images through a perfect foreground overlap', () => {
		const blank = new Uint8ClampedArray(8);
		expect(comparePixels(blank, blank).foregroundIntersectionOverUnion).toBe(0);
	});
	it('counts displaced content without translating it into alignment', () => {
		const a = new Uint8ClampedArray([0, 0, 0, 255, 255, 255, 255, 255]);
		const b = new Uint8ClampedArray([255, 255, 255, 255, 0, 0, 0, 255]);
		expect(comparePixels(a, b)).toMatchObject({
			normalizedMeanAbsoluteError: 1,
			changedPixelRatio: 1,
			foregroundIntersectionOverUnion: 0,
		});
	});
});
