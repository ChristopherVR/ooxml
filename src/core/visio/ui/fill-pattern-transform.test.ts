import { expect, it } from 'vitest';
import { composeVisioTransform, visioFillPatternTransform } from './fill-pattern-transform';

it('keeps page hatch axes fixed through composed rotation, scale and reflection', () => {
	for (const world of [
		[0, 1, -1, 0, 0, 0],
		[0, 2, -3, 0, 0, 0],
		[-1, 0, 0, 1, 0, 0],
	] as const) {
		expect(composeVisioTransform(world, visioFillPatternTransform(world)!)).toEqual([
			1, 0, 0, -1, 0, 0,
		]);
	}
	expect(visioFillPatternTransform([0, 0, 0, 1, 0, 0])).toBeUndefined();
	expect(visioFillPatternTransform([1, 0, 0, 1, 0, 0], NaN)).toBeUndefined();
});

it('retains the native page-height phase of the quarter-turned rectangle', () => {
	const world = [0, 1, -1, 0, 2.5, 0.5] as const;
	expect(composeVisioTransform(world, visioFillPatternTransform(world, 3)!)).toEqual([
		1, 0, 0, -1, -0.5, 0.5,
	]);
});
