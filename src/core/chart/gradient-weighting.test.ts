import { expect, it } from 'vitest';
import reference from './__fixtures__/native-gradient-two-stop-weighting.json';

// Excel 16.0 build 20430 captures of red (x%) to green (y%) two-stop fills with unequal and
// swapped alphas. They bound how far native colour weighting departs from straight
// (non-premultiplied) interpolation, which is what the renderer paints.
it('weights translucent two-stop colour almost like straight interpolation', () => {
	for (const sample of reference.cases.filter((c) => c.angle === 0)) {
		const leanSign = sample.profile === 'rg-unequal' ? 1 : -1;
		for (const pixel of sample.samples) {
			const t = (pixel.x + 0.5) / sample.width;
			const [r, g] = pixel.rgb as [number, number, number];
			const fraction = g / (r + g);
			// Premultiplied interpolation would lean by up to 0.08; native leans under 0.02 and
			// always toward the more opaque stop.
			expect((fraction - t) * leanSign, sample.name).toBeGreaterThan(-0.01);
			expect(Math.abs(fraction - t), sample.name).toBeLessThan(0.02);
		}
	}
});
