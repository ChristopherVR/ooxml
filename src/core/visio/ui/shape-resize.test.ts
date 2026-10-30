import { expect, it } from 'vitest';
import { demoDocument } from './demo-document';
import { transform } from '../geometry';
import { VISIO_RESIZE_HANDLES, visioResizeDrag, visioResizeShape } from './shape-resize';

function page(angle = 0, flipX = false, flipY = false) {
	const page = structuredClone(demoDocument.pages[0]!);
	page.connectors = [];
	page.shapes = [
		{
			...page.shapes[0]!,
			id: '1',
			width: 2,
			height: 1,
			rotation: { pinX: 4, pinY: 5, angle },
			transform: transform(4, 5, 0.25, 0.75, angle, flipX, flipY),
		},
	];
	return page;
}
it.each(VISIO_RESIZE_HANDLES)(
	'keeps the opposite $id anchor fixed through rotated/flipped resize',
	(handle) => {
		for (const [angle, flipX, flipY] of [
			[0, false, false],
			[Math.PI / 6, false, false],
			[Math.PI / 2, true, false],
			[-Math.PI / 4, false, true],
		] as const) {
			const scene = page(angle, flipX, flipY),
				shape = scene.shapes[0]!,
				before = structuredClone(shape);
			const [a, b, c, d, e, f] = shape.transform;
			const dx = handle.x === 0.5 ? 0 : handle.x === 1 ? 0.5 : -0.5,
				dy = handle.y === 0.5 ? 0 : handle.y === 1 ? 0.25 : -0.25;
			const result = visioResizeDrag(
				scene,
				'1',
				handle.id,
				{ x: 0, y: 0 },
				{ x: a * dx + c * dy, y: -(b * dx + d * dy) },
			)!;
			const { frame, command } = result,
				anchor = command.anchor!;
			const [na, nb, nc, nd, ne, nf] = frame.transform;
			expect(ne + na * anchor.x * frame.width + nc * anchor.y * frame.height).toBeCloseTo(
				e + a * anchor.x * shape.width + c * anchor.y * shape.height,
				12,
			);
			expect(nf + nb * anchor.x * frame.width + nd * anchor.y * frame.height).toBeCloseTo(
				f + b * anchor.x * shape.width + d * anchor.y * shape.height,
				12,
			);
			expect(frame.width).toBeCloseTo(handle.x === 0.5 ? 2 : 2.5);
			expect(frame.height).toBeCloseTo(handle.y === 0.5 ? 1 : 1.25);
			expect(frame.transform.slice(0, 4)).toEqual(shape.transform.slice(0, 4));
			expect(shape).toEqual(before);
		}
	},
);
it.each([0.5, 1, 2])(
	'emits drawing dimensions and unchanged normalized anchors at page scale%s',
	(scale) => {
		const scene = page();
		scene.drawingToPageScale = scale;
		expect(visioResizeDrag(scene, '1', 'ne', { x: 0, y: 0 }, { x: 1, y: -0.5 })!.command).toEqual({
			type: 'resize-shape',
			pageId: scene.id,
			shapeId: '1',
			width: 3 / scale,
			height: 1.5 / scale,
			anchor: { x: 0, y: 0 },
		});
	},
);
it('clamps crossing handles without changing flips and leaves tiny untouched dimensions unchanged', () => {
	const scene = page();
	const frame = visioResizeDrag(scene, '1', 'ne', { x: 0, y: 0 }, { x: -10, y: 10 })!.frame;
	// A corner resizes freely, so each side stops at 1/16 inch when it crosses the opposite one.
	expect(frame.width).toBe(1 / 16);
	expect(frame.height).toBe(1 / 16);
	const side = visioResizeDrag(scene, '1', 'e', { x: 0, y: 0 }, { x: -10, y: 10 })!.frame;
	expect(side.width).toBe(1 / 16);
	expect(side.height).toBe(1);
	scene.shapes[0]!.width = 0.01;
	scene.shapes[0]!.height = 0.01;
	const tiny = visioResizeDrag(scene, '1', 'ne', { x: 0, y: 0 }, { x: 0, y: 0 })!.frame;
	expect(tiny.width).toBe(0.01);
	expect(tiny.height).toBe(0.01);
});
it('declines unproven affine, grouped, glued, ambiguous and nonfinite candidates', () => {
	const scene = page();
	scene.shapes[0]!.transform = [2, 0, 0, 1, 0, 0];
	expect(visioResizeShape(scene, '1')).toBeUndefined();
	scene.shapes[0]!.transform = transform(4, 5, 0.25, 0.75, 0);
	expect(visioResizeDrag(scene, '1', 'ne', { x: 0, y: 0 }, { x: NaN, y: 1 })).toBeUndefined();
	scene.shapes.push(structuredClone(scene.shapes[0]!));
	expect(visioResizeShape(scene, '1')).toBeUndefined();
});

it('offers flat group handles with rotated parents and refuses unsupported member trees', () => {
	const scene = page(Math.PI / 6, true),
		group = scene.shapes[0]!;
	const member = structuredClone(page().shapes[0]!);
	member.id = '2';
	group.kind = 'group';
	group.children = [member];
	expect(visioResizeShape(scene, '1')).toBe(group);
	expect(visioResizeDrag(scene, '1', 'ne', { x: 0, y: 0 }, { x: 1, y: -1 })).toBeDefined();
	member.transform = transform(4, 5, 0.25, 0.75, 0.1);
	expect(visioResizeShape(scene, '1')).toBeUndefined();
	member.transform = transform(4, 5, 0.25, 0.75, 0);
	member.children = [structuredClone(page().shapes[0]!)];
	expect(visioResizeShape(scene, '1')).toBeUndefined();
	member.children = [];
	member.masterId = '0';
	expect(visioResizeShape(scene, '1')).toBeUndefined();
	delete member.masterId;
	member.layerIds = ['0'];
	expect(visioResizeShape(scene, '1')).toBeUndefined();
	delete member.layerIds;
	scene.connectors = [{ fromShapeId: '3', toShapeId: '2', fromCell: 'BeginX', toCell: 'PinX' }];
	expect(visioResizeShape(scene, '1')).toBeUndefined();
});

it('keeps proportions with Shift on corners, and on side handles, about the centre', () => {
	const scene = page();
	// Without Shift a corner resizes freely; with Shift the axis that changed more wins: +0.25 in
	// width is 12.5%, +0.5 in height is 50%.
	const free = visioResizeDrag(scene, '1', 'ne', { x: 0, y: 0 }, { x: 0.25, y: -0.5 })!;
	expect([free.frame.width, free.frame.height]).toEqual([2.25, 1.5]);
	const corner = visioResizeDrag(
		scene,
		'1',
		'ne',
		{ x: 0, y: 0 },
		{ x: 0.25, y: -0.5 },
		{ proportional: true },
	)!;
	expect(corner.frame.width).toBeCloseTo(3);
	expect(corner.frame.height).toBeCloseTo(1.5);
	expect(corner.command.anchor).toEqual({ x: 0, y: 0 });
	const sideFree = visioResizeDrag(scene, '1', 'e', { x: 0, y: 0 }, { x: 1, y: 0 })!;
	expect([sideFree.frame.width, sideFree.frame.height]).toEqual([3, 1]);
	const shift = visioResizeDrag(
		scene,
		'1',
		'e',
		{ x: 0, y: 0 },
		{ x: 1, y: 0 },
		{
			proportional: true,
		},
	)!;
	expect(shift.frame.width).toBeCloseTo(3);
	expect(shift.frame.height).toBeCloseTo(1.5);
	expect(shift.command.anchor).toEqual({ x: 0, y: 0.5 });
	const pin = scene.shapes[0]!.rotation!;
	// The left edge stays put and the vertical centre does not move.
	expect(shift.frame.transform[4]).toBeCloseTo(scene.shapes[0]!.transform[4]);
	expect(shift.frame.transform[5] + shift.frame.height / 2).toBeCloseTo(
		scene.shapes[0]!.transform[5] + scene.shapes[0]!.height / 2,
	);
	expect(pin).toEqual(page().shapes[0]!.rotation);
});
