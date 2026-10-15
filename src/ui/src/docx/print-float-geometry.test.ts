import { describe, expect, it } from 'vitest';
import { MIN_FLOAT_SIZE, moveBox, placementAt, resizeBox } from './print-float-geometry';

const start = { x: 100, y: 200, width: 200, height: 100 };

describe('float resize and move geometry', () => {
	it('keeps the aspect ratio and the opposite corner for corner handles', () => {
		const grown = resizeBox(start, 'se', 100, 0, true);
		expect(grown).toEqual({ x: 100, y: 200, width: 250, height: 125 });
		const shrunk = resizeBox(start, 'nw', 50, 25, true);
		expect(shrunk.width / shrunk.height).toBeCloseTo(2);
		// The bottom-right corner stays where it was.
		expect(shrunk.x + shrunk.width).toBeCloseTo(300);
		expect(shrunk.y + shrunk.height).toBeCloseTo(300);
	});

	it('averages both axes so a diagonal drag scales smoothly', () => {
		const box = resizeBox(start, 'ne', 40, -40, true);
		expect(box.width / box.height).toBeCloseTo(2);
		expect(box.width).toBeGreaterThan(200);
	});

	it('changes one dimension for edge handles and leaves the other alone', () => {
		expect(resizeBox(start, 'e', 30, 99, true)).toEqual({ ...start, width: 230 });
		expect(resizeBox(start, 'n', 99, -20, true)).toEqual({ ...start, y: 180, height: 120 });
		expect(resizeBox(start, 'w', 20, 0, false)).toEqual({ ...start, x: 120, width: 180 });
	});

	it('does not collapse below the minimum size', () => {
		const tiny = resizeBox(start, 'se', -500, -500, true);
		expect(tiny.width).toBeGreaterThanOrEqual(MIN_FLOAT_SIZE);
		expect(tiny.height).toBeGreaterThanOrEqual(MIN_FLOAT_SIZE);
		expect(resizeBox(start, 'e', -500, 0, false).width).toBe(MIN_FLOAT_SIZE);
	});

	it('moves without changing size', () => {
		expect(moveBox(start, -10, 5)).toEqual({ x: 90, y: 205, width: 200, height: 100 });
	});

	it('turns a box into offsets from the frame origin and drops alignment', () => {
		const placement = placementAt(
			{ wrap: 'square', align: 'right', alignV: 'bottom', relativeFrom: 'margin' },
			{ x: 340.4, y: 130.6, width: 10, height: 10 },
			{ x: 100, y: 100 },
		);
		expect(placement).toEqual({
			wrap: 'square',
			relativeFrom: 'margin',
			offsetXPx: 240,
			offsetYPx: 31,
		});
		expect(placementAt(undefined, start, { x: 0, y: 0 })).toEqual({
			wrap: 'none',
			offsetXPx: 100,
			offsetYPx: 200,
		});
	});
});
