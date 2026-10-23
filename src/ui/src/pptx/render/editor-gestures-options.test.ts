// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';

import type { GestureController, GestureDeps, GestureTransform } from './editor-gestures';
import { createGestureController } from './editor-gestures';
import type { InteractionBox } from './element-interaction';
import type { SnapLine } from './snap-guides';

/**
 * The deps the React and Angular canvases added when they moved onto the
 * shared controller: a mouse-down start (no pointer id), a kind-aware start
 * box, a minimum size, grid-snapped resize, a multi-selection snap exclusion,
 * a binding's own snap model and the mid-gesture lock re-check.
 */

const controllers: GestureController[] = [];

afterEach(() => {
	for (const controller of controllers.splice(0)) {
		controller.dispose();
	}
});

function setup(start: InteractionBox, overrides: Partial<GestureDeps> = {}) {
	const previews: Array<{ transform: GestureTransform; lines: readonly SnapLine[] }> = [];
	const ends: Array<{ transform: GestureTransform | null; moved: boolean }> = [];
	const kinds: string[] = [];
	const controller = createGestureController({
		getScale: () => 1,
		getElementBox: (_id, kind) => {
			kinds.push(kind);
			return start;
		},
		getSiblings: () => [],
		getStageOrigin: () => ({ left: 0, top: 0 }),
		onStart: () => undefined,
		onPreview: (transform, lines) => previews.push({ transform, lines }),
		onEnd: (transform, moved) => ends.push({ transform, moved }),
		...overrides,
	});
	controllers.push(controller);
	return { controller, previews, ends, kinds };
}

function move(type: 'pointermove' | 'pointerup', x: number, y: number, pointerId = 1): void {
	const event = new PointerEvent(type, { clientX: x, clientY: y });
	Object.defineProperty(event, 'pointerId', { value: pointerId, configurable: true });
	window.dispatchEvent(event);
}

const BOX = { x: 0, y: 0, width: 100, height: 50, rotation: 0 };

describe('createGestureController options', () => {
	it('follows any pointer when the press carried no pointer id (a mousedown)', () => {
		const { controller, previews } = setup(BOX);
		controller.begin('move', 'a', { clientX: 0, clientY: 0, shiftKey: false });
		move('pointermove', 10, 0, 7);
		expect(previews[0]?.transform.x).toBe(10);
	});

	it('tells getElementBox which kind of gesture is starting', () => {
		const { controller, kinds } = setup(BOX);
		controller.begin('resize', 'a', { clientX: 0, clientY: 0, shiftKey: false }, 'se');
		expect(kinds).toStrictEqual(['resize']);
	});

	it('clamps a resize to getMinSize', () => {
		const { controller, previews } = setup(BOX, { getMinSize: () => 30 });
		controller.begin('resize', 'a', { clientX: 0, clientY: 0, shiftKey: false }, 'se');
		move('pointermove', -200, -200);
		expect(previews[0]?.transform).toMatchObject({ width: 30, height: 30 });
	});

	it('snaps the dragged edges of a resize to the grid and leaves the anchor', () => {
		const { controller, previews } = setup(
			{ x: 3, y: 3, width: 100, height: 50 },
			{ getSnapToGrid: () => true, getGridSize: () => 10 },
		);
		controller.begin('resize', 'a', { clientX: 0, clientY: 0, shiftKey: false }, 'se');
		move('pointermove', 14, 4);
		// Right edge 117 -> 120, bottom 57 -> 60; the top-left anchor stays at 3,3.
		expect(previews[0]?.transform).toMatchObject({ x: 3, y: 3, width: 117, height: 57 });
	});

	it('ignores the rest of a dragged selection when snapping to shapes', () => {
		const sibling = { id: 'b', x: 113, y: 200, width: 20, height: 20 };
		const plain = setup(BOX, { getSiblings: () => [sibling] });
		plain.controller.begin('move', 'a', { clientX: 0, clientY: 0, shiftKey: false });
		move('pointermove', 10, 0);
		move('pointerup', 10, 0);
		expect(plain.previews[0]?.transform.x).toBe(13);

		const grouped = setup(BOX, {
			getSiblings: () => [sibling],
			getSnapExcludeIds: () => ['b'],
		});
		grouped.controller.begin('move', 'a', { clientX: 0, clientY: 0, shiftKey: false });
		move('pointermove', 10, 0);
		expect(grouped.previews[0]?.transform.x).toBe(10);
	});

	it('lets a binding replace the move snap with its own model', () => {
		const lines: SnapLine[] = [{ axis: 'v', position: 40 }];
		const { controller, previews } = setup(BOX, {
			getSnapToGrid: () => true,
			snapMove: (box) => ({ x: box.x + 1, y: 99, lines }),
		});
		controller.begin('move', 'a', { clientX: 0, clientY: 0, shiftKey: false });
		move('pointermove', 7, 0);
		expect(previews[0]?.transform).toMatchObject({ x: 8, y: 99 });
		expect(previews[0]?.lines).toBe(lines);
	});

	it('ends the gesture unmoved, with no further frames, once canContinue refuses', () => {
		let allowed = true;
		const { controller, previews, ends } = setup(BOX, { canContinue: () => allowed });
		controller.begin('move', 'a', { clientX: 0, clientY: 0, pointerId: 1, shiftKey: false });
		move('pointermove', 10, 0);
		allowed = false;
		move('pointermove', 20, 0);
		move('pointerup', 20, 0);
		expect(previews).toHaveLength(1);
		expect(ends).toStrictEqual([{ transform: null, moved: false }]);
		expect(controller.isActive()).toBeFalsy();
	});
});
