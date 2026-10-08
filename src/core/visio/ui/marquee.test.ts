import { expect, it } from 'vitest';
import { demoDocument } from './demo-document';
import { transform } from '../geometry';
import { visioMarqueeBox, visioMarqueeSelection, visioShapePageBox } from './marquee';

it('selects fully enclosed visible top-level shapes in page order, excluding partial and hidden content', () => {
	const page = structuredClone(demoDocument.pages[0]!);
	const template = page.shapes[0]!;
	page.shapes = ['a', 'b', 'c'].map((id, index) => ({
		...structuredClone(template),
		id,
		name: id,
		width: 2,
		height: 1,
		transform: transform(2 + index * 3, 5, 1, 0.5, 0),
	}));
	const box = visioMarqueeBox({ x: 7, y: page.height - 3 }, { x: 0, y: page.height - 7 })!;
	expect(visioMarqueeSelection(page, box)).toEqual(
		['a', 'b'].map((id) => ({ id, name: id, pageId: page.id })),
	);
	const visible = new WeakMap([[page.shapes[0]!, false]]);
	expect(visioMarqueeSelection(page, box, visible).map((item) => item.id)).toEqual(['b']);
	page.shapes[1]!.hidden = true;
	expect(visioMarqueeSelection(page, box)).toHaveLength(1);
});
it('uses all transformed corners for rotation, flips and noncentral pins', () => {
	const page = structuredClone(demoDocument.pages[0]!);
	const shape = page.shapes[0]!;
	shape.width = 4;
	shape.height = 1;
	shape.transform = transform(8, 2, 0.25, 0.5, Math.PI / 2);
	const box = visioShapePageBox(page, shape)!;
	expect(box.x).toBeCloseTo(7.5);
	expect(box.width).toBeCloseTo(1);
	expect(box.y).toBeCloseTo(page.height - 5.75);
	expect(box.height).toBeCloseTo(4);
	shape.transform = [-1, 0, 0, 1, 8, 2];
	expect(visioShapePageBox(page, shape)).toEqual({ x: 4, y: page.height - 3, width: 4, height: 1 });
});
it('allows read-only visible shapes to be selected without admitting source movement', () => {
	const page = structuredClone(demoDocument.pages[0]!);
	page.shapes = [page.shapes[0]!];
	page.shapes[0]!.masterId = '7';
	expect(visioMarqueeSelection(page, { x: -100, y: -100, width: 200, height: 200 })[0]?.id).toBe(
		page.shapes[0]!.id,
	);
	expect(visioMarqueeSelection(page, { x: 0, y: 0, width: 0, height: 20 })).toEqual([]);
	expect(visioMarqueeBox({ x: NaN, y: 0 }, { x: 1, y: 2 })).toBeUndefined();
	expect(
		visioMarqueeBox({ x: -Number.MAX_VALUE, y: 0 }, { x: Number.MAX_VALUE, y: 2 }),
	).toBeUndefined();
});
