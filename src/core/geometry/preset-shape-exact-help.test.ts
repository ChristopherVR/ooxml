/**
 * actionButtonHelp: the question mark is a chain of elliptical arcs fitted to
 * PowerPoint's raster of a 500 pt button (mean absolute difference 0.0014 over
 * the slide, 0.4 percent of the glyph's pixels off by half a level). The old
 * glyph was a coarse polyline of lines.
 */

import { describe, expect, it } from 'vitest';

import { evaluatePresetShape } from './preset-shape-evaluator';

describe('actionButtonHelp', () => {
	const paths = evaluatePresetShape('actionButtonHelp', 1000, 1000, undefined)!.paths;

	it('draws the question mark with the measured arcs', () => {
		const d = paths[1]!.d;
		expect(d.match(/ A /g)).toHaveLength(6);
		// Hook: outer radius .214 ss round (500, 338), flat cut at y = 338.
		expect(d).toContain('M 286 338 A 214 212.5');
		// Stem: width .106 ss, bottom .187 ss below the centre.
		expect(d).toContain('L 553 687 L 447 687');
	});

	it('draws the dot as a circle of radius .08 ss', () => {
		expect(paths[2]!.d).toContain('A 80 80');
		expect(paths[2]!.d.startsWith('M 580 795')).toBe(true);
	});
});
