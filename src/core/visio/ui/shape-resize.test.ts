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
		expect(visioResizeDrag(scene, '1', 'ne', { x: 0, y: 0 }, { x: 1, y: -1 })!.command).toEqual({
			type: 'resize-shape',
			pageId: scene.id,
			shapeId: '1',
			width: 3 / scale,
			height: 2 / scale,
			anchor: { x: 0, y: 0 },
		});
	},
);
it('clamps crossing handles without changing flips and leaves tiny untouched dimensions unchanged', () => {
	const scene = page();
	const frame = visioResizeDrag(scene, '1', 'ne', { x: 0, y: 0 }, { x: -10, y: 10 })!.frame;
	expect(frame.width).toBe(1 / 16);
	expect(frame.height).toBe(1 / 16);
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
