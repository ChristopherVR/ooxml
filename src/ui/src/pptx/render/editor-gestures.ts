import type { GestureSnap, GestureTransform } from './editor-gesture-frames';
import {
	computeMoveFrame,
	computeResizeFrame,
	computeRotateFrame,
	startRotationFrame,
	toGestureTransform,
} from './editor-gesture-frames';
import type { InteractionBox, Point, ResizeHandleId } from './element-interaction';
import type { SnapGuideInput, SnapLine, SnapSibling } from './snap-guides';

export type { GestureSnap, GestureTransform } from './editor-gesture-frames';

/**
 * Pointer gesture driver for the editing overlay: move / resize / rotate, used
 * by all five bindings.
 *
 * The per-frame maths is `editor-gesture-frames` (the shared
 * `element-interaction` helpers, `createRotationDrag` and the snap models);
 * this module only owns the pointer-event lifecycle (dead zone, window
 * listeners, cancel). A binding supplies what differs through its deps: the
 * boxes, the snap settings, the minimum size, the ids a multi-selection drag
 * ignores when snapping, its own snap model, a mid-gesture lock re-check, and
 * what a preview frame or the end of the gesture writes.
 *
 * The public surface (`begin`) takes a {@link PointerLike} plain object
 * instead of a real `PointerEvent`, so it is testable without constructing DOM
 * events. A `PointerEvent` (or a `MouseEvent`, which has no `pointerId`)
 * satisfies it structurally. `begin` never calls `preventDefault` /
 * `stopPropagation`; the caller's `pointerdown` handler does that.
 */

export type GestureKind = 'move' | 'resize' | 'rotate';

/** The subset of a `PointerEvent` this module needs. */
export interface PointerLike {
	clientX: number;
	clientY: number;
	/**
	 * Absent on a `mousedown` (a `MouseEvent`): the gesture then follows any
	 * pointer, as there is only one mouse.
	 */
	pointerId?: number;
	shiftKey: boolean;
}

export interface GestureDeps {
	/** Current stage scale (screen px per element px). */
	getScale(): number;
	/**
	 * Element geometry at gesture start (element px). `kind` lets a binding
	 * leave the rotation out of a move or resize box when it resizes axis-aligned.
	 */
	getElementBox(id: string, kind: GestureKind): InteractionBox | undefined;
	/** Sibling boxes on the same slide, for snap-to-shape during a move. */
	getSiblings(): SnapSibling[];
	getSnapToGrid?(): boolean;
	getSnapToShape?(): boolean;
	getGridSize?(): number;
	getGuides?(): readonly SnapGuideInput[];
	/** Smallest width/height a resize may reach (the geometry default when absent). */
	getMinSize?(): number;
	/** Ids the shape snap ignores besides `id` itself (the rest of a dragged selection). */
	getSnapExcludeIds?(id: string): Iterable<string>;
	/** A binding's own move snap, replacing the grid and shape snap. */
	snapMove?(box: InteractionBox, id: string): GestureSnap;
	/**
	 * Re-checked on every move: false (a lock arrived mid-gesture) ends the
	 * gesture where it is, as an unmoved one, with no further frames.
	 */
	canContinue?(id: string, kind: GestureKind): boolean;
	/** Overlay origin in client coordinates, for rotation pointer mapping. */
	getStageOrigin(): { left: number; top: number };
	/** First movement past the dead zone (push history, mark interaction). */
	onStart(id: string, kind: GestureKind): void;
	/** Live preview: apply the geometry and render `lines` as snap guides. */
	onPreview(transform: GestureTransform, lines: readonly SnapLine[]): void;
	/** Gesture finished. `moved` is false for a plain tap (no dead-zone exit). */
	onEnd(transform: GestureTransform | null, moved: boolean, id: string): void;
}

export interface GestureController {
	/** Begin a gesture from a press. `handle` only for `resize`. */
	begin(kind: GestureKind, id: string, pointer: PointerLike, handle?: ResizeHandleId): void;
	isActive(): boolean;
	/** Abort listeners without emitting an end transform (teardown). */
	dispose(): void;
}

/** Dead zone in screen px before a pointerdown becomes a drag. */
const DRAG_DEAD_ZONE_PX = 2;

interface ActiveGesture {
	kind: GestureKind;
	id: string;
	pointerId: number | undefined;
	startClientX: number;
	startClientY: number;
	startBox: InteractionBox;
	rotate?: (pointer: Point) => number;
	handle?: ResizeHandleId;
	moved: boolean;
	last: GestureTransform | null;
}

type Frame = readonly [GestureTransform, readonly SnapLine[]];

export function createGestureController(deps: GestureDeps): GestureController {
	let active: ActiveGesture | null = null;

	function stagePoint(pointer: PointerLike): Point {
		const scale = deps.getScale() || 1;
		const origin = deps.getStageOrigin();
		return {
			x: (pointer.clientX - origin.left) / scale,
			y: (pointer.clientY - origin.top) / scale,
		};
	}

	function frame(g: ActiveGesture, pointer: PointerLike, dx: number, dy: number): Frame {
		const scale = deps.getScale();
		const snapToGrid = deps.getSnapToGrid?.() ?? false;
		const gridSize = deps.getGridSize?.() ?? 12;
		if (g.kind === 'move') {
			return computeMoveFrame(g.id, g.startBox, dx, dy, {
				scale,
				snapToGrid,
				gridSize,
				snapToShape: deps.getSnapToShape?.() !== false,
				siblings: deps.getSiblings(),
				excludeIds: new Set([g.id, ...(deps.getSnapExcludeIds?.(g.id) ?? [])]),
				guides: deps.getGuides?.() ?? [],
				...(deps.snapMove ? { snapMove: deps.snapMove } : {}),
			});
		}
		if (g.kind === 'resize') {
			const minSize = deps.getMinSize?.();
			const next = computeResizeFrame(g.id, g.startBox, g.handle, dx, dy, {
				scale,
				lockAspect: pointer.shiftKey,
				gridSize: snapToGrid ? gridSize : null,
				...(minSize === undefined ? {} : { minSize }),
			});
			return [next, []];
		}
		const point = stagePoint(pointer);
		return [computeRotateFrame(g.id, g.startBox, g.rotate, point, pointer.shiftKey), []];
	}

	const samePointer = (g: ActiveGesture, pointer: PointerLike): boolean =>
		g.pointerId === undefined || pointer.pointerId === g.pointerId;

	function finish(g: ActiveGesture, transform: GestureTransform | null, moved: boolean): void {
		detach();
		active = null;
		deps.onEnd(transform, moved, g.id);
	}

	function onPointerMove(pointer: PointerLike): void {
		const g = active;
		if (!g || !samePointer(g, pointer)) {
			return;
		}
		if (deps.canContinue && !deps.canContinue(g.id, g.kind)) {
			finish(g, null, false);
			return;
		}
		const dx = pointer.clientX - g.startClientX;
		const dy = pointer.clientY - g.startClientY;
		if (!g.moved) {
			if (Math.abs(dx) <= DRAG_DEAD_ZONE_PX && Math.abs(dy) <= DRAG_DEAD_ZONE_PX) {
				return;
			}
			g.moved = true;
			deps.onStart(g.id, g.kind);
		}
		const [next, lines] = frame(g, pointer, dx, dy);
		g.last = next;
		deps.onPreview(next, lines);
	}

	function onPointerUp(pointer: PointerLike): void {
		const g = active;
		if (!g || !samePointer(g, pointer)) {
			return;
		}
		finish(g, g.moved ? g.last : null, g.moved);
	}

	function detach(): void {
		window.removeEventListener('pointermove', onPointerMove);
		window.removeEventListener('pointerup', onPointerUp);
		window.removeEventListener('pointercancel', onPointerCancel);
	}

	function onPointerCancel(pointer: PointerLike): void {
		const g = active;
		if (!g || !samePointer(g, pointer)) {
			return;
		}
		// Repaint the original geometry without committing the interrupted preview.
		if (g.moved) {
			deps.onPreview(toGestureTransform(g.id, g.startBox), []);
		}
		finish(g, null, false);
	}

	return {
		begin(kind, id, pointer, handle) {
			const box = deps.getElementBox(id, kind);
			if (!box || active) {
				return;
			}
			const startBox = { ...box };
			active = {
				kind,
				id,
				pointerId: pointer.pointerId,
				startClientX: pointer.clientX,
				startClientY: pointer.clientY,
				startBox,
				moved: false,
				last: null,
			};
			if (kind === 'rotate') {
				active.rotate = startRotationFrame(startBox, stagePoint(pointer));
			}
			if (handle) {
				active.handle = handle;
			}
			window.addEventListener('pointermove', onPointerMove);
			window.addEventListener('pointerup', onPointerUp);
			window.addEventListener('pointercancel', onPointerCancel);
		},
		isActive() {
			return active !== null;
		},
		dispose() {
			detach();
			active = null;
		},
	};
}
