import { describe, expect, it } from 'vitest';

import { evaluatePresetShape } from './preset-shape-evaluator';

/** The route of a connector preset in a 200 x 150 box, as numbers. */
function route(name: string, adjustments?: Record<string, number>): number[] {
	const d = evaluatePresetShape(name, 200, 150, adjustments)?.svgPath ?? '';
	return (d.match(/-?\d+(?:\.\d+)?/gu) ?? []).map(Number);
}

/**
 * The control points of ECMA-376's `curvedConnector3/4/5`, which PowerPoint 16
 * draws (a static export of each preset, at default and edited adjustments,
 * differs from these curves by under 0.05 % of the pixels; the earlier
 * hand-written approximations differed by about 0.5 %, a visibly different
 * S whose middle was not vertical).
 */
describe('curved connector presets follow ECMA-376', () => {
	it.each([
		['curvedConnector2', undefined, [0, 0, 100, 0, 200, 75, 200, 150]],
		[
			'curvedConnector3',
			undefined,
			[0, 0, 50, 0, 100, 37.5, 100, 75, 100, 112.5, 150, 150, 200, 150],
		],
		[
			'curvedConnector3',
			{ adj1: 25000 },
			[0, 0, 25, 0, 50, 37.5, 50, 75, 50, 112.5, 125, 150, 200, 150],
		],
		[
			'curvedConnector4',
			{ adj1: 25000, adj2: 75000 },
			[
				0, 0, 25, 0, 50, 28.125, 50, 56.25, 50, 84.375, 87.5, 112.5, 125, 112.5, 162.5, 112.5, 200,
				131.25, 200, 150,
			],
		],
		[
			'curvedConnector5',
			{ adj1: 20000, adj2: 30000, adj3: 80000 },
			[
				0, 0, 20, 0, 40, 11.25, 40, 22.5, 40, 33.75, 70, 45, 100, 45, 130, 45, 160, 71.25, 160,
				97.5, 160, 123.75, 180, 150, 200, 150,
			],
		],
	])('%s %j', (name, adjustments, expected) => {
		const actual = route(name, adjustments);
		expect(actual).toHaveLength(expected.length);
		expected.forEach((value, i) => expect(actual[i]).toBeCloseTo(value, 2));
	});
});
