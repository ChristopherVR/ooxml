import { expect, it } from 'vitest';
import { buildChartGradientDef } from './gradient-definition';
it('retains the native vertical vector and resolves off-box circle focus', () => {
	expect(buildChartGradientDef('g', { type: 'linear', angle: 90, stops: [] })).toMatchObject({
		x1: 0.5,
		y1: 0,
		x2: 0.5,
		y2: 1,
	});
	expect(
		buildChartGradientDef('g', { type: 'radial', focalPoint: { x: 0.5, y: -0.8 }, stops: [] }),
	).toMatchObject({ cx: 0.5, cy: -0.8, r: Math.hypot(0.5, 1.8) });
});
