import { describe, expect, it } from 'vitest';

import { fitGlowChain, fitGlowUnderGhost } from './morph-effect-glow';

describe('morph glow dissolve fit', () => {
	it('starts and ends on the two glows alone', () => {
		const a = { radius: 20, alpha: 0.75 };
		const b = { radius: 9, alpha: 0.75 };
		// `b` is the narrower, so it is painted first: [b alpha, a alpha].
		const [n0, w0] = fitGlowChain(a, b, 0);
		expect(n0).toBeCloseTo(0, 3);
		expect(w0).toBeCloseTo(0.75, 3);
		const [n1, w1] = fitGlowChain(a, b, 1);
		expect(n1).toBeCloseTo(0.75, 3);
		expect(w1).toBeCloseTo(0, 3);
	});

	it('takes alpha off the wide glow and puts it on the narrow one mid-dissolve', () => {
		// The wide glow painted second also glows the narrow one, so the fit
		// keeps it below its plain fade, and lifts the narrow one above its own.
		const [narrow, wide] = fitGlowChain(
			{ radius: 9, alpha: 0.75 },
			{ radius: 20, alpha: 0.75 },
			0.5,
		);
		expect(wide).toBeLessThan(0.375);
		expect(narrow).toBeGreaterThan(0.375);
		expect(narrow).toBeLessThan(0.75);
	});

	it('fades the incoming glow in under a ghost faster than p', () => {
		const alpha = fitGlowUnderGhost({ radius: 20, alpha: 1 }, { radius: 20, alpha: 1 }, 0.5);
		expect(alpha).toBeGreaterThan(0.5);
		expect(alpha).toBeLessThanOrEqual(1);
		expect(fitGlowUnderGhost({ radius: 20, alpha: 1 }, { radius: 20, alpha: 1 }, 0)).toBe(0);
	});
});
