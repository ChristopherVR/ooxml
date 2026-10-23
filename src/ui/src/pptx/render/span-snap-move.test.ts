import { describe, expect, it } from 'vitest';

import { snapMoveToSpanGuides } from './span-snap-move';

const box = { x: 103, y: 50, width: 40, height: 20 };
const other = { x: 100, y: 200, width: 60, height: 30 };
const off = { others: [], threshold: 6, snapToShape: false, gridSize: null, rulerGuides: null };

describe('snapMoveToSpanGuides', () => {
	it('leaves the box alone with every snap off', () => {
		expect(snapMoveToSpanGuides(box, off)).toStrictEqual({ x: 103, y: 50, guides: [] });
	});

	it('snaps to the closest edge of another box and draws a span guide', () => {
		const result = snapMoveToSpanGuides(box, { ...off, others: [other], snapToShape: true });
		expect(result.x).toBe(100);
		expect(result.guides).toContainEqual(expect.objectContaining({ axis: 'x', pos: 100 }));
	});

	it('applies the grid after the shape snap', () => {
		const result = snapMoveToSpanGuides(box, {
			...off,
			others: [other],
			snapToShape: true,
			gridSize: 16,
		});
		expect(result).toMatchObject({ x: 96, y: 48 });
	});

	it('snaps an edge or the centre to a ruler guide last', () => {
		const result = snapMoveToSpanGuides(box, {
			...off,
			rulerGuides: [
				{ axis: 'x', pos: 125 },
				{ axis: 'y', pos: 72 },
			],
		});
		// The centre (123) lands on x=125; the bottom edge (70) lands on y=72.
		expect(result).toMatchObject({ x: 105, y: 52 });
	});
});
