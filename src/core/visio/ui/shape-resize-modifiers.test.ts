import { describe, expect, it } from 'vitest';
import { demoDocument } from './demo-document';
import { transform } from '../geometry';
import type { VisioPage } from '../model';
import {
	visioResizeDrag,
	type VisioResizeDragOptions,
	type VisioResizeHandle,
} from './shape-resize';

/** One 2 x 1 inch shape pinned at (4, 5), with its local origin at (3.75, 4.25) when unrotated. */
function page(angle = 0, locks: VisioPage['shapes'][number]['resizeLocks'] = undefined): VisioPage {
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

/** Page drag in y-up inches, as the viewer converts a pointer move into start and end points. */
function drag(dx: number, dy: number) {
	return { start: { x: 0, y: 0 }, end: { x: dx, y: -dy } };
}

function resize(
	scene: VisioPage,
	handle: VisioResizeHandle,
	dx: number,
	dy: number,
	options: VisioResizeDragOptions = {},
) {
	const { start, end } = drag(dx, dy);
	const result = visioResizeDrag(scene, '1', handle, start, end, options);
	expect(result).toBeDefined();
	return result!;
}

/** The page position of the frame's centre, which Ctrl resizing keeps fixed. */
function centre(frame: { transform: number[]; width: number; height: number }) {
	const [a, b, c, d, e, f] = frame.transform as [number, number, number, number, number, number];
	return {
		x: e + a * (frame.width / 2) + c * (frame.height / 2),
		y: f + b * (frame.width / 2) + d * (frame.height / 2),
	};
}

describe('corner handles resize freely unless LockAspect is set', () => {
	it('moves the two edges independently with no modifier', () => {
		const { frame } = resize(page(), 'ne', 1, 0.25);
		expect(frame.width).toBeCloseTo(3, 12);
		expect(frame.height).toBeCloseTo(1.25, 12);
	});

	it('keeps the original proportions with Shift, scaling by the axis that changed more', () => {
		const { frame, command } = resize(page(), 'ne', 1, 0.25, { proportional: true });
		expect(frame.width).toBeCloseTo(3, 12);
		expect(frame.height).toBeCloseTo(1.5, 12);
		expect(command.anchor).toEqual({ x: 0, y: 0 });
	});

	it('keeps the proportions on a LockAspect shape without Shift', () => {
		const { frame } = resize(page(0, { aspect: true }), 'ne', 1, 0.25);
		expect(frame.width).toBeCloseTo(3, 12);
		expect(frame.height).toBeCloseTo(1.5, 12);
	});

	it('keeps the proportions on a LockAspect shape with a side handle, about its centre', () => {
		const scene = page(0, { aspect: true });
		const before = centre(resize(scene, 'e', 0, 0).frame);
		const { frame } = resize(scene, 'e', 1, 0);
		expect(frame.width).toBeCloseTo(3, 12);
		expect(frame.height).toBeCloseTo(1.5, 12);
		const after = centre(frame);
		expect(after.x).toBeCloseTo(before.x + 0.5, 12);
		expect(after.y).toBeCloseTo(before.y, 12);
	});
});

describe('side handles change one dimension', () => {
	it('changes only the width with no modifier', () => {
		const { frame } = resize(page(), 'e', 1, 0);
		expect(frame.width).toBeCloseTo(3, 12);
		expect(frame.height).toBe(1);
	});

	it('keeps the aspect ratio about the centre with Shift', () => {
		const { frame } = resize(page(), 'e', 1, 0, { proportional: true });
		expect(frame.width).toBeCloseTo(3, 12);
		expect(frame.height).toBeCloseTo(1.5, 12);
	});
});

describe('Ctrl resizes about the centre', () => {
	it('grows both edges of a corner by the drag, keeping the centre fixed', () => {
		const scene = page();
		const before = centre(resize(scene, 'ne', 0, 0).frame);
		const { frame, command } = resize(scene, 'ne', 1, 0.25, { centre: true });
		expect(frame.width).toBeCloseTo(4, 12);
		expect(frame.height).toBeCloseTo(1.5, 12);
		expect(command.anchor).toEqual({ x: 0.5, y: 0.5 });
		const after = centre(frame);
		expect(after.x).toBeCloseTo(before.x, 12);
		expect(after.y).toBeCloseTo(before.y, 12);
	});

	it('moves only the width of a side handle, with its centre fixed', () => {
		const scene = page();
		const before = centre(resize(scene, 'w', 0, 0).frame);
		const { frame } = resize(scene, 'w', -0.5, 0, { centre: true });
		expect(frame.width).toBeCloseTo(3, 12);
		expect(frame.height).toBe(1);
		const after = centre(frame);
		expect(after.x).toBeCloseTo(before.x, 12);
	});

	it('keeps the aspect ratio with Ctrl and Shift together', () => {
		const scene = page();
		const before = centre(resize(scene, 'ne', 0, 0).frame);
		const { frame } = resize(scene, 'ne', 1, 0.25, { centre: true, proportional: true });
		// Both edges move, so the width grows by 2 and the height scales by the same 2.
		expect(frame.width).toBeCloseTo(4, 12);
		expect(frame.height).toBeCloseTo(2, 12);
		const after = centre(frame);
		expect(after.x).toBeCloseTo(before.x, 12);
		expect(after.y).toBeCloseTo(before.y, 12);
	});

	it('keeps the centre fixed through a LockAspect corner with Ctrl', () => {
		const scene = page(0, { aspect: true });
		const before = centre(resize(scene, 'sw', 0, 0).frame);
		const { frame } = resize(scene, 'sw', -0.5, -0.5, { centre: true });
		expect(frame.width / frame.height).toBeCloseTo(2, 12);
		const after = centre(frame);
		expect(after.x).toBeCloseTo(before.x, 12);
		expect(after.y).toBeCloseTo(before.y, 12);
	});
});

describe('resize locks', () => {
	it('holds the width of a LockWidth shape on corner drags, leaving the height free', () => {
		const { frame } = resize(page(0, { width: true }), 'ne', 1, 0.25);
		expect(frame.width).toBe(2);
		expect(frame.height).toBeCloseTo(1.25, 12);
	});

	it('holds a LockWidth dimension against a side drag', () => {
		const { frame } = resize(page(0, { width: true }), 'e', 1, 0);
		expect(frame.width).toBe(2);
		expect(frame.height).toBe(1);
	});

	it('holds the height of a LockHeight shape on a top handle', () => {
		const { frame } = resize(page(0, { height: true }), 'n', 0, 1);
		expect(frame.width).toBe(2);
		expect(frame.height).toBe(1);
	});

	it('leaves a LockHeight shape free to widen on its side handle', () => {
		const { frame } = resize(page(0, { height: true }), 'e', 1, 0);
		expect(frame.width).toBeCloseTo(3, 12);
		expect(frame.height).toBe(1);
	});

	it('does not move a LockAspect shape with a locked width under Shift', () => {
		const { frame } = resize(page(0, { aspect: true, width: true }), 'ne', 1, 0.25, {
			proportional: true,
		});
		expect(frame.width).toBe(2);
		expect(frame.height).toBe(1);
	});
});

describe('rotated shapes resize in their local frame', () => {
	// A 90 degree turn: page +x is local -y, so page +y widens the local width.
	const angle = Math.PI / 2;

	it('widens the local width on a free corner drag', () => {
		const { frame } = resize(page(angle), 'ne', 0, 0.5);
		expect(frame.width).toBeCloseTo(2.5, 12);
		expect(frame.height).toBe(1);
	});

	it('keeps the local proportions with Shift', () => {
		const { frame } = resize(page(angle), 'ne', 0, 0.5, { proportional: true });
		expect(frame.width).toBeCloseTo(2.5, 12);
		expect(frame.height).toBeCloseTo(1.25, 12);
	});

	it('resizes about the page-space centre with Ctrl', () => {
		const scene = page(angle);
		const before = centre(resize(scene, 'ne', 0, 0).frame);
		const { frame } = resize(scene, 'ne', 0, 0.5, { centre: true });
		expect(frame.width).toBeCloseTo(3, 12);
		expect(frame.height).toBe(1);
		const after = centre(frame);
		expect(after.x).toBeCloseTo(before.x, 12);
		expect(after.y).toBeCloseTo(before.y, 12);
	});

	it('applies a LockWidth lock to the local width, not the page width', () => {
		const { frame } = resize(page(angle, { width: true }), 'ne', 0, 0.5);
		expect(frame.width).toBe(2);
		expect(frame.height).toBe(1);
	});
});
