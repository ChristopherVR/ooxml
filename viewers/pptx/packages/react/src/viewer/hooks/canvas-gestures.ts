import { createGestureController, publishLiveGeometry } from 'ooxml-ui/pptx';
import type {
	GestureController,
	GestureKind,
	GestureTransform,
	PointerLike,
	ResizeHandleId,
	SnapLine,
} from 'ooxml-ui/pptx';
/**
 * The canvas's move and resize gestures on the shared `createGestureController`
 * (the one the other four bindings drive). The controller owns the dead zone,
 * the window pointer listeners and the per-frame maths, snap-to-grid and
 * snap-to-shape included; this keeps what is React's own:
 *
 *   - frames are written straight to the DOM (no React state per frame), at
 *     most once per animation frame, and mirrored to collaborators;
 *   - a drag moves every movable member of the selection by the primary's
 *     snapped delta, and the shape snap ignores all of them;
 *   - the commit lands once, on release, through `updateSlides` (or the
 *     template store), rerouting the connectors glued to what moved.
 *
 * `dragStateRef` / `resizeStateRef` stay the record of the armed gesture (the
 * rest of the viewer reads them as "a pointer interaction is running"); the
 * press handlers fill them and then call `begin`.
 */
import type { PptxElement } from 'pptx-viewer-core';

import { MIN_ELEMENT_SIZE } from '../constants';
import type { DragState, ResizeState } from '../types';
import {
	applyReroutedConnectors,
	rerouteConnectorsForMovedElements,
} from '../utils/connector-reroute';
import { syncSelectionHandleOverlay } from '../utils/selection-handle-overlay';
import type { UsePointerHandlersInput } from './pointer-handler-types';

export interface CanvasGestures {
	/** Start the gesture armed in `dragStateRef` (move) or `resizeStateRef` (resize). */
	begin(kind: Exclude<GestureKind, 'rotate'>, id: string, pointer: PointerLike): void;
	isActive(): boolean;
	dispose(): void;
}

/** Schedules one preview write; `requestAnimationFrame` in the browser. */
export interface FrameScheduler {
	request(callback: () => void): number;
	cancel(handle: number): void;
}

const animationFrames: FrameScheduler = {
	request: (callback) => requestAnimationFrame(callback),
	cancel: (handle) => cancelAnimationFrame(handle),
};

type Armed = { kind: 'move'; drag: DragState } | { kind: 'resize'; resize: ResizeState };

export function createCanvasGestures(
	getInput: () => UsePointerHandlersInput,
	frames: FrameScheduler = animationFrames,
): CanvasGestures {
	let armed: Armed | null = null;
	let pending: { transform: GestureTransform; lines: readonly SnapLine[] } | null = null;
	let frameHandle = 0;
	let lastSnapLinesKey = '';

	/** Where interim geometry is mirrored, or undefined when it must not be. */
	const liveTarget = () => {
		const { livePatcher, editTemplateMode, activeSlide } = getInput();
		return livePatcher && !editTemplateMode && livePatcher.isActive()
			? { patcher: livePatcher, slideId: activeSlide?.id }
			: undefined;
	};

	function showSnapLines(lines: readonly SnapLine[]): void {
		const next = lines.map((line) => ({
			axis: line.axis === 'v' ? ('x' as const) : ('y' as const),
			position: line.position,
		}));
		const key = JSON.stringify(next);
		if (key !== lastSnapLinesKey) {
			lastSnapLinesKey = key;
			getInput().setSnapLines(next);
		}
	}

	function paintMove(drag: DragState, transform: GestureTransform): void {
		const start = drag.startPositionsById[transform.id];
		if (!start) {
			return;
		}
		drag.lastDx = transform.x - start.x;
		drag.lastDy = transform.y - start.y;
		const live = liveTarget();
		for (const [id, from] of Object.entries(drag.startPositionsById)) {
			const x = from.x + drag.lastDx;
			const y = from.y + drag.lastDy;
			const domEl = drag.domEls.get(id);
			if (domEl) {
				domEl.style.left = `${x}px`;
				domEl.style.top = `${y}px`;
				syncSelectionHandleOverlay(domEl, id, { x, y });
			}
			// The DOM writes bypass React state, so this is all peers see before the commit.
			if (live) {
				publishLiveGeometry(live.patcher, live.slideId, id, { x, y });
			}
		}
	}

	function paintResize(resize: ResizeState, transform: GestureTransform): void {
		const { x, y } = transform;
		const width = Math.max(transform.width, MIN_ELEMENT_SIZE);
		const height = Math.max(transform.height, MIN_ELEMENT_SIZE);
		resize.lastX = x;
		resize.lastY = y;
		resize.lastWidth = transform.width;
		resize.lastHeight = transform.height;
		if (resize.domEl) {
			Object.assign(resize.domEl.style, {
				left: `${x}px`,
				top: `${y}px`,
				width: `${width}px`,
				height: `${height}px`,
			});
			syncSelectionHandleOverlay(resize.domEl, resize.elementId, { x, y, width, height });
		}
		const live = liveTarget();
		if (live) {
			publishLiveGeometry(live.patcher, live.slideId, resize.elementId, { x, y, width, height });
		}
	}

	function paintPending(): void {
		const frame = pending;
		pending = null;
		frameHandle = 0;
		if (!frame || !armed) {
			return;
		}
		if (armed.kind === 'move') {
			armed.drag.moved = true;
			showSnapLines(frame.lines);
			paintMove(armed.drag, frame.transform);
		} else {
			armed.resize.moved = true;
			paintResize(armed.resize, frame.transform);
		}
	}

	function flushFrame(): void {
		if (frameHandle !== 0) {
			frames.cancel(frameHandle);
		}
		paintPending();
	}

	const controller: GestureController = createGestureController({
		getScale: () => getInput().editorScale,
		// Axis-aligned, as the canvas has always moved and resized (no rotation).
		getElementBox: (id) => {
			const el = getInput().elementLookup.get(id);
			return el ? { x: el.x, y: el.y, width: el.width, height: el.height } : undefined;
		},
		getSiblings: () => {
			const { editTemplateMode, templateElements, activeSlide } = getInput();
			const source = editTemplateMode ? templateElements : (activeSlide?.elements ?? []);
			return source.map(({ id, x, y, width, height }) => ({ id, x, y, width, height }));
		},
		getSnapToGrid: () => getInput().snapToGrid,
		getGridSize: () => getInput().gridSpacingPx,
		getSnapToShape: () => getInput().snapToShape,
		getGuides: () => getInput().guides.map(({ axis, position }) => ({ axis, position })),
		getMinSize: () => MIN_ELEMENT_SIZE,
		getSnapExcludeIds: () =>
			armed?.kind === 'move' ? Object.keys(armed.drag.startPositionsById) : [],
		getStageOrigin: () => {
			const rect = getInput().canvasStageRef.current?.getBoundingClientRect();
			return { left: rect?.left ?? 0, top: rect?.top ?? 0 };
		},
		onStart: () => undefined,
		onPreview: (transform, lines) => {
			pending = { transform, lines };
			if (frameHandle === 0) {
				frameHandle = frames.request(paintPending);
			}
		},
		onEnd: (transform, moved) => {
			flushFrame();
			const ended = armed;
			armed = null;
			if (ended) {
				finishGesture(ended, moved ? transform : null);
			}
		},
	});

	function finishGesture(ended: Armed, transform: GestureTransform | null): void {
		const input = getInput();
		if (transform) {
			if (ended.kind === 'move') {
				commitMove(input, ended.drag);
			} else {
				commitResize(input, ended.resize);
			}
			// The pointer ends back over the node it went down on, so the browser
			// still fires a trailing `click` there: flag it so the click handler does
			// not read it as a second click on a selected element (inline edit).
			input.justInteractedRef.current = true;
			// Push the last interim frame now so peers land on the final geometry.
			input.livePatcher?.flush();
		}
		input.dragStateRef.current = null;
		input.resizeStateRef.current = null;
		lastSnapLinesKey = '';
		input.setSnapLines([]);
		if (transform) {
			input.markDirty();
		}
		// Pointer-down can commit text before arming the gesture; wake history even
		// for a still pointer (its document comparison rejects no-op clicks).
		input.setPointerCommitNonce((n) => n + 1);
	}

	return {
		begin(kind, id, pointer) {
			if (controller.isActive()) {
				return;
			}
			const input = getInput();
			const drag = input.dragStateRef.current;
			const resize = input.resizeStateRef.current;
			armed =
				kind === 'move' && drag
					? { kind, drag }
					: kind === 'resize' && resize
						? { kind, resize }
						: null;
			if (!armed) {
				return;
			}
			lastSnapLinesKey = '';
			const handle: ResizeHandleId | undefined =
				armed.kind === 'resize' ? armed.resize.handle : undefined;
			controller.begin(kind, id, pointer, handle);
			if (!controller.isActive()) {
				armed = null;
			}
		},
		isActive: () => controller.isActive(),
		dispose: () => {
			if (frameHandle !== 0) {
				frames.cancel(frameHandle);
				frameHandle = 0;
			}
			pending = null;
			armed = null;
			controller.dispose();
		},
	};
}

/** Move every dragged element by the committed delta and reroute their connectors. */
function commitMove(input: UsePointerHandlersInput, drag: DragState): void {
	const { lastDx: dx, lastDy: dy } = drag;
	const movedIds = new Set(Object.keys(drag.startPositionsById));
	const moveElements = (elements: PptxElement[]): PptxElement[] => {
		const moved = elements.map((el) => {
			const start = drag.startPositionsById[el.id];
			return start ? { ...el, x: start.x + dx, y: start.y + dy } : el;
		});
		return applyReroutedConnectors(moved, rerouteConnectorsForMovedElements(moved, movedIds));
	};
	if (input.editTemplateMode) {
		// A template (master/layout) element commits to the template store;
		// buildSaveSlides merges it back so the edit persists to the shared part.
		const slideId = input.activeSlide?.id;
		if (slideId) {
			input.setTemplateElementsBySlideId((prev) => ({
				...prev,
				[slideId]: moveElements(prev[slideId] ?? []),
			}));
		}
		return;
	}
	input.updateSlides((prev) =>
		prev.map((s, i) =>
			i !== input.activeSlideIndex ? s : { ...s, elements: moveElements(s.elements) },
		),
	);
}

/** Apply the resize and reroute the connectors glued to the resized element. */
function commitResize(input: UsePointerHandlersInput, resize: ResizeState): void {
	input.updateElementById(resize.elementId, {
		x: resize.lastX,
		y: resize.lastY,
		width: resize.lastWidth,
		height: resize.lastHeight,
	});
	const movedIds = new Set([resize.elementId]);
	input.updateSlides((prev) =>
		prev.map((s, i) => {
			if (i !== input.activeSlideIndex) {
				return s;
			}
			const rerouted = rerouteConnectorsForMovedElements(s.elements, movedIds);
			return rerouted.length === 0
				? s
				: { ...s, elements: applyReroutedConnectors(s.elements, rerouted) };
		}),
	);
}
