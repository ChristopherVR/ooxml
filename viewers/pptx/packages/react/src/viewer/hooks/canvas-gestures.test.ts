// @vitest-environment happy-dom
/**
 * The React canvas's move and resize on the shared gesture controller: a
 * multi-selection drag moves every member by the primary's snapped delta and
 * snaps against the other shapes only, the release commits once (flagging the
 * trailing click), a cancelled drag repaints where it started without a
 * commit, and a resize snaps its dragged edges to the grid.
 */
import { createCollaborationLivePatcher } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createCanvasGestures } from './canvas-gestures';
import type { CanvasGestures, FrameScheduler } from './canvas-gestures';
import type { UsePointerHandlersInput } from './pointer-handler-types';

const box = (id: string, x: number, y: number, width = 100, height = 50): PptxElement =>
	({ id, type: 'shape', shapeType: 'rect', x, y, width, height }) as unknown as PptxElement;

const syncFrames: FrameScheduler = {
	request: (callback) => {
		callback();
		return 1;
	},
	cancel: () => undefined,
};

const running: CanvasGestures[] = [];
afterEach(() => {
	for (const gestures of running.splice(0)) {
		gestures.dispose();
	}
	document.body.replaceChildren();
});

function makeInput(
	elements: PptxElement[],
	overrides: Partial<UsePointerHandlersInput> = {},
): UsePointerHandlersInput {
	const activeSlide = { id: 's1', slideNumber: 1, elements } as unknown as PptxSlide;
	return {
		editorScale: 1,
		canvasStageRef: { current: null },
		canvasSize: { width: 960, height: 540 },
		activeSlide,
		activeSlideIndex: 0,
		gridSpacingPx: 10,
		dragStateRef: { current: null },
		resizeStateRef: { current: null },
		shapeAdjustmentDragStateRef: { current: null },
		marqueeStateRef: { current: null },
		justInteractedRef: { current: false },
		editTemplateMode: false,
		snapToGrid: false,
		snapToShape: true,
		guides: [],
		templateElements: [],
		elementLookup: new Map(elements.map((el) => [el.id, el])),
		setMarqueeSelectionState: vi.fn(),
		setSnapLines: vi.fn(),
		setTemplateElementsBySlideId: vi.fn(),
		setPointerCommitNonce: vi.fn(),
		effectiveSelectedIds: [],
		applySelection: vi.fn(),
		clearSelection: vi.fn(),
		updateSlides: vi.fn(),
		updateElementById: vi.fn(),
		markDirty: vi.fn(),
		livePatcher: createCollaborationLivePatcher(),
		...overrides,
	} as UsePointerHandlersInput;
}

function start(input: UsePointerHandlersInput, kind: 'move' | 'resize', id: string): void {
	const gestures = createCanvasGestures(() => input, syncFrames);
	running.push(gestures);
	gestures.begin(kind, id, { clientX: 0, clientY: 0, shiftKey: false });
}

const fire = (type: 'pointermove' | 'pointerup' | 'pointercancel', x: number, y: number) =>
	window.dispatchEvent(new PointerEvent(type, { clientX: x, clientY: y }));

function armDrag(input: UsePointerHandlersInput, ids: string[]): Map<string, HTMLElement> {
	const domEls = new Map<string, HTMLElement>();
	const startPositionsById: Record<string, { x: number; y: number }> = {};
	for (const id of ids) {
		const el = input.elementLookup.get(id)!;
		startPositionsById[id] = { x: el.x, y: el.y };
		const node = document.createElement('div');
		document.body.appendChild(node);
		domEls.set(id, node);
	}
	input.dragStateRef.current = {
		elementId: ids[0],
		startClientX: 0,
		startClientY: 0,
		startPositionsById,
		domEls,
		moved: false,
		lastDx: 0,
		lastDy: 0,
	};
	return domEls;
}

describe('canvas gestures', () => {
	it('drags a multi-selection by the primary snapped delta, snapping to other shapes only', () => {
		// b sits where a would snap if it were a sibling; it is dragged too, so
		// only c (whose left edge is at 213) may pull a's right edge.
		const input = makeInput([box('a', 0, 0), box('b', 108, 0), box('c', 213, 300)]);
		const domEls = armDrag(input, ['a', 'b']);
		start(input, 'move', 'a');
		fire('pointermove', 110, 0);
		expect(domEls.get('a')?.style.left).toBe('113px');
		expect(domEls.get('b')?.style.left).toBe('221px');
		expect(input.setSnapLines).toHaveBeenLastCalledWith([{ axis: 'x', position: 213 }]);
		expect(input.updateSlides).not.toHaveBeenCalled();

		fire('pointerup', 110, 0);
		expect(input.updateSlides).toHaveBeenCalledOnce();
		const updater = vi.mocked(input.updateSlides).mock.calls[0][0];
		const [slide] = updater([input.activeSlide!]);
		expect(slide.elements.map((el) => el.x)).toStrictEqual([113, 221, 213]);
		expect(input.justInteractedRef.current).toBeTruthy();
		expect(input.markDirty).toHaveBeenCalledOnce();
		expect(input.dragStateRef.current).toBeNull();
		expect(input.setSnapLines).toHaveBeenLastCalledWith([]);
	});

	it('wakes history for a still press without committing or flagging a click', () => {
		const input = makeInput([box('a', 0, 0)]);
		armDrag(input, ['a']);
		start(input, 'move', 'a');
		fire('pointerup', 1, 1);
		expect(input.updateSlides).not.toHaveBeenCalled();
		expect(input.justInteractedRef.current).toBeFalsy();
		expect(input.setPointerCommitNonce).toHaveBeenCalledOnce();
		expect(input.dragStateRef.current).toBeNull();
	});

	it('repaints a cancelled drag where it started and commits nothing', () => {
		const input = makeInput([box('a', 40, 20)], { snapToShape: false });
		const domEls = armDrag(input, ['a']);
		start(input, 'move', 'a');
		fire('pointermove', 30, 30);
		expect(domEls.get('a')?.style.left).toBe('70px');
		fire('pointercancel', 30, 30);
		expect(domEls.get('a')?.style.left).toBe('40px');
		expect(input.updateSlides).not.toHaveBeenCalled();
		expect(input.markDirty).not.toHaveBeenCalled();
	});

	it('resizes with grid-snapped dragged edges and commits the box once', () => {
		const input = makeInput([box('a', 3, 3)], { snapToGrid: true });
		input.resizeStateRef.current = {
			elementId: 'a',
			startClientX: 0,
			startClientY: 0,
			startX: 3,
			startY: 3,
			startWidth: 100,
			startHeight: 50,
			handle: 'se',
			moved: false,
			domEl: null,
			lastX: 3,
			lastY: 3,
			lastWidth: 100,
			lastHeight: 50,
		};
		start(input, 'resize', 'a');
		fire('pointermove', 14, 4);
		fire('pointerup', 14, 4);
		expect(input.updateElementById).toHaveBeenCalledExactlyOnceWith('a', {
			x: 3,
			y: 3,
			width: 117,
			height: 57,
		});
		expect(input.resizeStateRef.current).toBeNull();
	});
});
