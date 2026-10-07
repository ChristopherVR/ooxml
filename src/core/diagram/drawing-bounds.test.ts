import { describe, expect, it } from 'vitest';

import { computeDiagramDrawingBounds } from './drawing-bounds.js';

describe('cached drawing bounds', () => {
	it('fits shape frames with negative coordinates', () => {
		expect(
			computeDiagramDrawingBounds([
				{ frame: { x: -50, y: -20, width: 30, height: 10 } },
				{ frame: { x: 20, y: 40, width: 10, height: 10 } },
			]),
		).toEqual({ minX: -50, minY: -20, width: 80, height: 70 });
	});

	it('includes text extending outside its shape in either direction', () => {
		expect(
			computeDiagramDrawingBounds([
				{
					frame: { x: 10, y: 20, width: 100, height: 50 },
					textFrame: { x: -20, y: 5, width: 200, height: 100 },
				},
			]),
		).toEqual({ minX: -20, minY: 5, width: 200, height: 100 });
	});

	it('does not convert the caller coordinate units', () => {
		expect(
			computeDiagramDrawingBounds([
				{
					frame: { x: 914400, y: 1828800, width: 914400, height: 457200 },
				},
			]),
		).toEqual({ minX: 914400, minY: 1828800, width: 914400, height: 457200 });
	});

	it('retains a unit fallback for empty and zero-area drawings', () => {
		expect(computeDiagramDrawingBounds([])).toEqual({ minX: 0, minY: 0, width: 1, height: 1 });
		expect(computeDiagramDrawingBounds([{ frame: { x: 5, y: 6, width: 0, height: 0 } }])).toEqual({
			minX: 5,
			minY: 6,
			width: 1,
			height: 1,
		});
	});
});
