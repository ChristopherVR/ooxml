import { expect, it } from 'vitest';
import { demoDocument } from './demo-document';
import {
	visioDrawPlan,
	visioDrawingPoint,
	visioBoxCreationCommand,
	visioDrawBounds,
} from './draw-plan';

it.each([0.5, 1, 2])('plans fixed bounds and native pins at page scale %s', (scale) => {
	const page = { ...demoDocument.pages[0]!, shapes: [], drawingToPageScale: scale };
	for (const kind of ['rectangle', 'ellipse', 'text'] as const) {
		const command = visioDrawPlan(page, kind, { x: 4, y: 3 }, { x: 2, y: 1 }, 'A\nB')!;
		expect(command).toMatchObject({
			shapeId: '1',
			x: 3 / scale,
			y: (page.height - 2) / scale,
			width: 2 / scale,
			height: 2 / scale,
			text: 'A\nB',
		});
	}
	expect(visioDrawPlan(page, 'line', { x: 1, y: 2 }, { x: 4, y: 3 })).toMatchObject({
		beginX: 1 / scale,
		beginY: (page.height - 2) / scale,
		endX: 4 / scale,
		endY: (page.height - 3) / scale,
	});
});
it('rejects invalid runtime tools, overflow, bad scale and exhausted IDs without returning a fallback shape', () => {
	const page = structuredClone(demoDocument.pages[0]!);
	const point = { x: 1, y: 2 },
		end = { x: 3, y: 4 };
	expect(visioDrawPlan(page, 'wrong' as 'text', point, end)).toBeUndefined();
	expect(() =>
		visioBoxCreationCommand(page, 'wrong' as 'text', point, { width: 1, height: 1 }),
	).toThrow('tool');
	expect(
		visioDrawPlan(page, 'text', { x: -Number.MAX_VALUE, y: 0 }, { x: Number.MAX_VALUE, y: 1 }),
	).toBeUndefined();
	expect(() => visioDrawBounds({ x: -Number.MAX_VALUE, y: 0 }, end)).toThrow('bounds');
	expect(() =>
		visioDrawPlan({ ...page, drawingToPageScale: Number.MIN_VALUE }, 'text', point, end),
	).toThrow('coordinate');
	page.shapes[0]!.id = '4294967295';
	expect(() => visioDrawPlan(page, 'text', point, end)).toThrow('IDs');
});
it('snaps physical coordinates and admits diagonal lines but refuses tiny boxes and invalid points', () => {
	const page = demoDocument.pages[0]!;
	expect(visioDrawingPoint(page, { x: 1.04, y: -1 })).toEqual({ x: 1.0625, y: 0 });
	expect(visioDrawingPoint(page, { x: 1.04, y: -1 }, { snap: false, bounded: false })).toEqual({
		x: 1.04,
		y: -1,
	});
	expect(visioDrawingPoint(page, { x: NaN, y: 0 })).toBeUndefined();
	expect(visioDrawPlan(page, 'text', { x: 0, y: 0 }, { x: 0.01, y: 1 })).toBeUndefined();
	expect(visioDrawPlan(page, 'line', { x: 0, y: 0 }, { x: 0.05, y: 0.05 })).toBeDefined();
	expect(visioDrawPlan(page, 'line', { x: 0, y: 0 }, { x: Infinity, y: 0 })).toBeUndefined();
});
