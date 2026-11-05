import { describe, expect, it } from 'vitest';

import {
	glowCssLengths,
	glowFilterLengths,
	shadowBlurToCss,
	SHADOW_BLUR_TO_CSS,
} from './effect-profile';

/** A CSS blur length `L` is a Gaussian of deviation `L / 2`. */
const deviation = (cssLength: number) => cssLength / 2;

describe('effect profiles measured on PowerPoint', () => {
	it('maps blurRad to a Gaussian of deviation 0.326 blurRad', () => {
		// Fitted edge deviations for blurRad 10, 20, 30, 50 pt at 1280 px: 4.33, 8.74,
		// 12.86 and 21.90 px, of blurRad 13.3, 26.7, 40 and 66.7 px.
		for (const [blur, fitted] of [
			[13.33, 4.33],
			[26.67, 8.74],
			[40, 12.86],
			[66.67, 21.9],
		] as const) {
			expect(deviation(shadowBlurToCss(blur))).toBeCloseTo(fitted, 0);
		}
		// the previous mapping (halving) painted a deviation of only 0.25 blurRad
		expect(SHADOW_BLUR_TO_CSS / 2).toBeGreaterThan(0.25 * 1.2);
	});

	it('maps a glow to a spread of 0.417 R and a deviation of R / 6', () => {
		// Fitted for rad 5, 10, 20 and 40 pt: centre 2.18, 5.16, 11.16, 22.17 px and
		// deviation 1.07, 2.19, 4.42, 8.89 px, of R = 6.67, 13.3, 26.7, 53.3 px.
		for (const [radius, centre, sigma] of [
			[6.67, 2.18, 1.07],
			[13.33, 5.16, 2.19],
			[26.67, 11.16, 4.42],
			[53.33, 22.17, 8.89],
		] as const) {
			const { spread, blur } = glowCssLengths(radius);
			expect(Math.abs(spread - centre)).toBeLessThan(0.7);
			expect(deviation(blur)).toBeCloseTo(sigma, 0);
		}
	});

	it('chains three equal drop-shadows for a glow that has no box to spread', () => {
		expect(glowFilterLengths(20)).toEqual([7.6, 7.6, 7.6]);
		expect(glowFilterLengths(-4)).toEqual([0, 0, 0]);
	});
});
