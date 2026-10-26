import { expect, it } from 'vitest';
import type { PptxSmartArtData, SmartArtLayoutType } from '../types';
import { reflowSmartArtLayout } from './smartart-editing-reflow';

const families: SmartArtLayoutType[] = [
	'list',
	'process',
	'hierarchy',
	'cycle',
	'matrix',
	'pyramid',
	'funnel',
	'target',
	'gear',
	'venn',
	'timeline',
	'relationship',
	'chevron',
	'bending',
];
const data = (family: SmartArtLayoutType): PptxSmartArtData => ({
	resolvedLayoutType: family,
	nodes: [
		{ id: 'a', text: 'Plan' },
		{ id: 'b', text: 'Build' },
		{ id: 'c', text: 'Ship' },
	],
});

it.each(families)(
	'uses a finite local origin when %s reflow receives only a frame size',
	(family) => {
		const shapes = reflowSmartArtLayout(data(family), { width: 733, height: 667 })!;
		expect(shapes.length).toBeGreaterThan(0);
		for (const shape of shapes)
			expect([shape.x, shape.y, shape.width, shape.height].every(Number.isFinite)).toBe(true);
		expect(shapes).toEqual(
			reflowSmartArtLayout(data(family), { x: 0, y: 0, width: 733, height: 667 }),
		);
	},
);

it('retains a supplied origin coordinate while defaulting only the omitted coordinate', () => {
	const origin = reflowSmartArtLayout(data('cycle'), { width: 733, height: 667 })!;
	const shifted = reflowSmartArtLayout(data('cycle'), { x: 23, width: 733, height: 667 })!;
	shifted.forEach((shape, index) => {
		expect(shape.x).toBeCloseTo(origin[index]!.x + 23);
		expect(shape.y).toBe(origin[index]!.y);
	});
});
