import { describe, expect, it } from 'vitest';

import { evaluatePresetShape } from './preset-shape-evaluator';

/**
 * The regular `pentagon` preset, against PowerPoint 16 (`Slide.Export`,
 * 1280 x 720) of a 300 x 240 pt pentagon: its side vertices sit 38% of the way
 * down the box (sin 18deg above the ellipse centre `svc`), the bottom pair on
 * the box bottom, and the path starts at the left vertex (ECMA-376's order,
 * which a morph between two `!!`-named shapes walks).
 */
describe('pentagon preset', () => {
	it('puts the side vertices 38% down and the bottom pair on the bottom', () => {
		const d = evaluatePresetShape('pentagon', 100, 50)?.svgPath ?? '';
		const points = [...d.matchAll(/(-?[\d.e-]+) (-?[\d.e-]+)/gu)].map((m) => [
			Math.round(Number(m[1]) * 100) / 100 + 0,
			Math.round(Number(m[2]) * 100) / 100 + 0,
		]);
		expect(points).toStrictEqual([
			[0, 19.1],
			[50, 0],
			[100, 19.1],
			[80.9, 50],
			[19.1, 50],
		]);
	});
});
