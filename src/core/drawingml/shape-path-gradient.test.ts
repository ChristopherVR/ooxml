import { describe, expect, it } from 'vitest';

import {
	buildOutlinePathGradientSvg,
	closedSubpaths,
	scalePathAbout,
	shapePathGradientField,
} from './shape-path-gradient';

const WHITE_TO_BLACK = [
	{ position: 0, color: '#ffffff' },
	{ position: 100, color: '#000000' },
];

describe('shapePathGradientField', () => {
	it('paints a freeform and the block arrows as the circle PowerPoint draws for them', () => {
		expect(shapePathGradientField(undefined)).toBe('circle');
		expect(shapePathGradientField('custom')).toBe('circle');
		expect(shapePathGradientField('rightArrow')).toBe('circle');
		expect(shapePathGradientField('swooshArrow')).toBe('circle');
		expect(shapePathGradientField('sun')).toBe('circle');
	});

	it('scales the outline of everything else, and keeps rings on the approximation', () => {
		expect(shapePathGradientField('ellipse')).toBe('outline');
		expect(shapePathGradientField('star5')).toBe('outline');
		expect(shapePathGradientField('roundRect')).toBe('outline');
		expect(shapePathGradientField('donut')).toBe('legacy');
		expect(shapePathGradientField('frame')).toBe('legacy');
	});
});

describe('scalePathAbout', () => {
	it('scales points about the focus and radii about nothing', () => {
		expect(scalePathAbout('M 0 50 A 50 50 0 0 1 50 0 L 100 100 Z', 0.5, 50, 50)).toBe(
			'M 25 50 A 25 25 0 0 1 50 25 L 75 75 Z',
		);
	});

	it('refuses relative commands', () => {
		expect(scalePathAbout('m 0 0 l 10 10 z', 0.5, 0, 0)).toBeUndefined();
	});
});

describe('closedSubpaths', () => {
	it('drops the open stroke-only lines the evaluator appends', () => {
		expect(closedSubpaths('M 0 0 L 10 0 L 10 10 Z M 5 0 L 5 10')).toBe('M 0 0 L 10 0 L 10 10 Z');
		expect(closedSubpaths('M 0 0 L 10 0')).toBeUndefined();
	});
});

describe('buildOutlinePathGradientSvg', () => {
	const d = 'M 0 0 L 100 0 L 100 100 L 0 100 Z';

	it('nests one copy per band, largest first, scaled about the focus', () => {
		const svg = buildOutlinePathGradientSvg(WHITE_TO_BLACK, d, 100, 100, 0.25, 0.75)!;
		const uses = [...svg.matchAll(/matrix\(([^)]*)\)/gu)].map((m) => m[1]!.split(' ').map(Number));
		expect(uses).toHaveLength(100);
		expect(uses[0]).toEqual([1, 0, 0, 1, 0, 0]);
		// s = 0.5: the focus (25, 75) stays put
		expect(uses[50]).toEqual([0.5, 0, 0, 0.5, 12.5, 37.5]);
		expect(svg).toContain('viewBox="0 0 100 100"');
	});

	it('colours each copy at the middle of its band', () => {
		const svg = buildOutlinePathGradientSvg(WHITE_TO_BLACK, d, 100, 100, 0.5, 0.5)!;
		const fills = [...svg.matchAll(/<use [^>]*fill="(#[0-9a-f]{6})"/gu)].map((m) => m[1]);
		// outermost band centre is position 99.5 (almost black), innermost 0.5
		expect(fills[0]).toBe('#010101');
		expect(fills[99]).toBe('#fefefe');
	});

	it('draws a translucent ramp as even-odd rings', () => {
		const svg = buildOutlinePathGradientSvg(
			[
				{ position: 0, color: '#ff0000', opacity: 1 },
				{ position: 100, color: '#ff0000', opacity: 0 },
			],
			d,
			100,
			100,
			0.5,
			0.5,
		)!;
		expect(svg).not.toContain('<use');
		expect(svg.match(/fill-rule="evenodd"/gu)).toHaveLength(100);
		expect(svg).toContain('fill-opacity=');
	});

	it('gives up on an outline it cannot scale', () => {
		expect(buildOutlinePathGradientSvg(WHITE_TO_BLACK, 'm 0 0 l 1 1 z', 10, 10, 0.5, 0.5)).toBe(
			undefined,
		);
	});
});
