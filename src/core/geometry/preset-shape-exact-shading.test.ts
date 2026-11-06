/**
 * Sub-path shading modes, measured on PowerPoint 16 with a mid-grey fill
 * (can lid 222, cube top 211 / side 161 / front 200, bevel left 222 / top 211 /
 * right 120 / bottom 161 / well 200).
 */

import { describe, expect, it } from 'vitest';

import { evaluatePresetShape } from './preset-shape-evaluator';

const modes = (name: string) =>
	evaluatePresetShape(name, 150, 90)!
		.paths.map((p) => p.fill ?? 'norm')
		.filter((m) => m !== 'none' && m !== 'norm');

describe('shading modes', () => {
	it('shades a cube lightenLess on top and darkenLess on the side, front plain', () => {
		expect(modes('cube')).toEqual(['darkenLess', 'lightenLess']);
	});

	it('shades a bevel lighten, darken, lightenLess, darkenLess and leaves the well plain', () => {
		expect(modes('bevel')).toEqual(['lighten', 'darken', 'lightenLess', 'darkenLess']);
	});

	it('lightens the can lid', () => {
		expect(modes('can')).toEqual(['lighten']);
	});
});
