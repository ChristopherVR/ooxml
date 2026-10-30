import { describe, expect, it } from 'vitest';
import { demoDocument } from './demo-document';
import { transform } from '../geometry';
import type { VisioPage, VisioShape } from '../model';
import { visioResizeDrag, type VisioResizeDragOptions } from './shape-resize';
import { visioResizeSnapEnd } from './shape-resize-snap';

const SNAP = { shapes: true, guides: true, threshold: 0.3 };

/**
 * A 2 x 1 inch shape pinned at (4, 5). Unrotated, its local origin is (3.75, 4.25), so its box is
 * x 3.75 to 5.75 and y 4.25 to 5.25 in page inches.
 */
function page(angle = 0, locks?: VisioShape['resizeLocks']): VisioPage {
	const scene = structuredClone(demoDocument.pages[0]!);
	scene.connectors = [];
	scene.shapes = [
		{
			...scene.shapes[0]!,
			id: '1',
			width: 2,
			height: 1,
			rotation: { pinX: 4, pinY: 5, angle },
			transform: transform(4, 5, 0.25, 0.75, angle),
			...(locks ? { resizeLocks: locks } : {}),
		},
	];
	return scene;
}

/** An unrotated shape of the given size whose lower-left corner is at page (x, y). */
function box(id: string, x: number, y: number, width = 1, height = 1): VisioShape {
	const shape = structuredClone(demoDocument.pages[0]!.shapes[0]!);
	return {
		...shape,
		id,
		width,
		height,
		transform: [1, 0, 0, 1, x, y],
		hidden: false,
		children: [],
		layerIds: undefined,
		rotation: { pinX: x + width / 2, pinY: y + height / 2, angle: 0 },
	} as VisioShape;
}

/**
 * A snap target at y 0 to 1, clear of the moving shape's y range in these tests, so that only the
 * x edge can snap to it.
 */
function sideTarget(id: string, x: number, width = 1): VisioShape {
	return box(id, x, 0, width, 1);
}

function snapped(
	scene: VisioPage,
	handle: 'e' | 'n' | 'ne' | 'w',
	dx: number,
	dy: number,
	options: VisioResizeDragOptions & { snap?: Partial<typeof SNAP> } = {},
) {
	const { snap = {}, ...modifiers } = options;
	return visioResizeSnapEnd(
		scene,
		'1',
		handle,
		{ x: 0, y: 0 },
		{ x: dx, y: 0 - dy },
		{ ...modifiers, snap: { ...SNAP, ...snap } },
	);
}

describe('snapping a resize edge to other shapes and guides', () => {
	it("moves the right edge onto another shape's left edge and shows the guide", () => {
		const scene = page();
		scene.shapes.push(sideTarget('2', 6));
		// The right edge lands at 5.95, 0.05 short of the other shape's left edge at 6.
		const result = snapped(scene, 'e', 0.2, 0);
		expect(result.end.x).toBeCloseTo(0.25, 12);
		expect(result.end.y).toBe(0);
		expect(result.lines).toEqual([expect.objectContaining({ axis: 'x', pos: 6 })]);
		const { frame } = visioResizeDrag(scene, '1', 'e', { x: 0, y: 0 }, result.end)!;
		expect(frame.width).toBeCloseTo(2.25, 12);
	});

	it('snaps to the centre of another shape when that is the nearest line', () => {
		const scene = page();
		// Its centre is at x 5.8; the right edge is at 5.85 after a drag of 0.1, nearest to 5.8.
		scene.shapes.push(sideTarget('2', 5.6, 0.4));
		const result = snapped(scene, 'e', 0.1, 0);
		expect(result.end.x).toBeCloseTo(0.05, 12);
		expect(result.lines).toEqual([expect.objectContaining({ axis: 'x', pos: 5.8 })]);
	});

	it('snaps a horizontal edge of a 90 degree turned shape onto page y', () => {
		const scene = page(Math.PI / 2);
		const shape = scene.shapes[0]!;
		const [, b, , d, , f] = shape.transform;
		// The top of the turned shape: its local x edge, which the 'e' handle moves.
		const top = b * shape.width + d * shape.height + f;
		// A shape whose top is 0.05 above that edge, so the edge lifts onto it.
		scene.shapes.push(box('2', 8, top + 0.05 - 1));
		const result = snapped(scene, 'e', 0, 0);
		expect(result.end.y).toBeCloseTo(-0.05, 12);
		expect(result.end.x).toBe(0);
		expect(result.lines).toEqual([expect.objectContaining({ axis: 'y' })]);
	});

	it('snaps to a ruler guide the same way as to a shape edge', () => {
		const scene = page();
		scene.shapes.push({
			...box('3', 6, 0, 0, 0),
			rotation: { pinX: 6, pinY: 0, angle: Math.PI / 2 },
			visibility: { layerHidden: false, guide: true, layerPrintSummary: 'unlayered' },
		});
		const result = snapped(scene, 'e', 0.2, 0, { snap: { shapes: false } });
		expect(result.end.x).toBeCloseTo(0.25, 12);
		expect(result.lines).toEqual([expect.objectContaining({ axis: 'x', pos: 6 })]);
	});

	it('leaves the handle alone when nothing is within the snap distance', () => {
		const scene = page();
		scene.shapes.push(sideTarget('2', 6));
		expect(snapped(scene, 'e', 0.2, 0, { snap: { threshold: 0.01 } })).toEqual({
			end: { x: 0.2, y: 0 },
			lines: [],
		});
	});

	it('does nothing when neither shapes nor guides are snapped to', () => {
		const scene = page();
		scene.shapes.push(sideTarget('2', 6));
		expect(snapped(scene, 'e', 0.2, 0, { snap: { shapes: false, guides: false } })).toEqual({
			end: { x: 0.2, y: 0 },
			lines: [],
		});
	});

	it('does not snap a proportional drag, which would move the other edge off its line', () => {
		const scene = page();
		scene.shapes.push(sideTarget('2', 6));
		expect(snapped(scene, 'e', 0.2, 0, { proportional: true }).lines).toEqual([]);
		// A corner keeps the proportions by default, so it does not snap either.
		expect(snapped(scene, 'ne', 0.2, 0).lines).toEqual([]);
	});

	it('does not snap a LockAspect shape', () => {
		const scene = page(0, { aspect: true });
		scene.shapes.push(sideTarget('2', 6));
		expect(snapped(scene, 'e', 0.2, 0).lines).toEqual([]);
	});

	it('does not snap the x edge of a LockWidth shape, but still snaps a top edge', () => {
		const locked = page(0, { width: true });
		locked.shapes.push(sideTarget('2', 6));
		expect(snapped(locked, 'e', 0.2, 0).lines).toEqual([]);
		const top = page(0, { width: true });
		top.shapes.push(box('4', 0, 5.3, 1, 1));
		const result = snapped(top, 'n', 0, 0.2);
		expect(result.lines.map((line) => line.axis)).toEqual(['y']);
		expect(result.end.x).toBe(0);
		// The top would stop at 5.45, so it snaps back down to the target's bottom at 5.3.
		expect(result.end.y).toBeCloseTo(-0.05, 12);
	});

	it('keeps the centre fixed with Ctrl and still snaps the handle onto the line', () => {
		const scene = page();
		scene.shapes.push(sideTarget('2', 6));
		const result = snapped(scene, 'e', 0.2, 0, { centre: true });
		expect(result.end.x).toBeCloseTo(0.25, 12);
		const { frame, command } = visioResizeDrag(scene, '1', 'e', { x: 0, y: 0 }, result.end, {
			centre: true,
		})!;
		// Each edge moves by the snapped drag, so the width grows twice as much.
		expect(frame.width).toBeCloseTo(2.5, 12);
		expect(command.anchor).toEqual({ x: 0.5, y: 0.5 });
	});

	it('does not snap a rotated edge that is off both page axes', () => {
		const scene = page(Math.PI / 4);
		scene.shapes.push(sideTarget('2', 6));
		expect(snapped(scene, 'e', 0.2, 0).lines).toEqual([]);
	});
});
