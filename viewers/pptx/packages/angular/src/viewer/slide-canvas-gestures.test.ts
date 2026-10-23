// @vitest-environment jsdom
/**
 * slide-canvas-gestures.test.ts: the Angular canvas's move / resize / rotate
 * on the shared gesture controller. Pins what the canvas did on its own before:
 * the span-guide move snap (shapes, then grid), the axis-aligned resize clamped
 * to MIN_RESIZE, whole-degree rotation, the mid-gesture lock gate, and the
 * reroute request after a move or resize but not after a rotate.
 */
import type { GestureKind, InteractionBox, SnapGuide } from 'ooxml-ui/pptx';
import { afterEach, describe, expect, it } from 'vitest';

import type { Box } from './drag-resize';
import { MIN_RESIZE } from './drag-resize';
import { createSlideCanvasGestures } from './slide-canvas-gestures';
import type { SlideCanvasGestureHost, SlideCanvasGestures } from './slide-canvas-gestures';

interface Log {
	starts: Array<{ id: string; kind: GestureKind }>;
	updates: Array<{ id: string; box: Box }>;
	rotations: number[];
	ends: Array<readonly string[]>;
	guides: Array<readonly SnapGuide[]>;
}

const live: SlideCanvasGestures[] = [];
afterEach(() => {
	for (const gestures of live.splice(0)) {
		gestures.dispose();
	}
});

function setup(
	box: InteractionBox,
	overrides: Partial<SlideCanvasGestureHost> = {},
): { gestures: SlideCanvasGestures; log: Log } {
	const log: Log = { starts: [], updates: [], rotations: [], ends: [], guides: [] };
	const gestures = createSlideCanvasGestures({
		scale: () => 1,
		stage: () => undefined,
		elementBox: () => box,
		canInteract: () => true,
		snapOthers: () => [],
		snapToShape: () => false,
		snapToGrid: () => false,
		gridSize: () => 10,
		rulerGuides: () => null,
		setSnapGuides: (guides) => log.guides.push(guides),
		start: (id, kind) => log.starts.push({ id, kind }),
		update: (id, next) => log.updates.push({ id, box: next }),
		rotate: (_id, rotation) => log.rotations.push(rotation),
		end: (ids) => log.ends.push(ids),
		...overrides,
	});
	live.push(gestures);
	return { gestures, log };
}

function fire(type: 'pointermove' | 'pointerup', x: number, y: number, shiftKey = false): void {
	const event = new PointerEvent(type, { clientX: x, clientY: y, shiftKey });
	Object.defineProperty(event, 'pointerId', { value: 1, configurable: true });
	window.dispatchEvent(event);
}

const press = { clientX: 0, clientY: 0, pointerId: 1, shiftKey: false };

describe('slide canvas gestures', () => {
	it('moves, snaps to another shape with span guides, and asks for a reroute', () => {
		const { gestures, log } = setup(
			{ x: 0, y: 0, width: 40, height: 20 },
			{
				snapToShape: () => true,
				snapOthers: () => [{ x: 100, y: 200, width: 60, height: 30 }],
			},
		);
		gestures.begin('move', 'a', press);
		fire('pointermove', 98, 0);
		expect(log.starts).toStrictEqual([{ id: 'a', kind: 'move' }]);
		expect(log.updates[0]?.box).toStrictEqual({ x: 100, y: 0, width: 40, height: 20 });
		expect(log.guides[0]).toContainEqual(expect.objectContaining({ axis: 'x', pos: 100 }));
		expect(gestures.moved()).toBeTruthy();
		fire('pointerup', 98, 0);
		expect(log.ends).toStrictEqual([['a']]);
		expect(log.guides.at(-1)).toStrictEqual([]);
		expect(gestures.moved()).toBeFalsy();
	});

	it('resizes axis-aligned, clamped to MIN_RESIZE, with grid-snapped edges', () => {
		const { gestures, log } = setup({ x: 3, y: 3, width: 100, height: 50 });
		gestures.begin('resize', 'a', press, 'se');
		fire('pointermove', -500, -500);
		expect(log.updates[0]?.box).toStrictEqual({
			x: 3,
			y: 3,
			width: MIN_RESIZE,
			height: MIN_RESIZE,
		});

		const snapped = setup({ x: 3, y: 3, width: 100, height: 50 }, { snapToGrid: () => true });
		snapped.gestures.begin('resize', 'a', press, 'se');
		fire('pointermove', 14, 4);
		expect(snapped.log.updates[0]?.box).toStrictEqual({ x: 3, y: 3, width: 117, height: 57 });
	});

	it('rotates in whole degrees and does not ask for a reroute', () => {
		const { gestures, log } = setup({ x: 0, y: 0, width: 100, height: 100, rotation: 0 });
		gestures.begin('rotate', 'a', { ...press, clientX: 50, clientY: 0 });
		fire('pointermove', 200, 53);
		expect(log.rotations[0]).toBe(Math.round(log.rotations[0] ?? 0.5));
		expect(log.rotations[0]).toBeGreaterThan(80);
		fire('pointermove', 200, 53, true);
		expect(log.rotations[1]).toBe(90);
		fire('pointerup', 200, 53);
		expect(log.ends).toStrictEqual([]);
	});

	it('stops writing frames once a lock forbids the gesture mid-drag', () => {
		let allowed = true;
		const { gestures, log } = setup(
			{ x: 0, y: 0, width: 40, height: 20 },
			{ canInteract: () => allowed },
		);
		gestures.begin('move', 'a', press);
		fire('pointermove', 10, 0);
		allowed = false;
		fire('pointermove', 30, 0);
		fire('pointerup', 30, 0);
		expect(log.updates).toHaveLength(1);
		expect(log.ends).toStrictEqual([]);
		expect(gestures.isActive()).toBeFalsy();
	});
});
