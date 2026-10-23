/**
 * slide-canvas-gestures.ts: the move / resize / rotate gestures of
 * `SlideCanvasComponent`, on the shared `createGestureController` that the
 * other four bindings drive too. The controller owns the dead zone, the window
 * pointer listeners and the per-frame maths; this maps the canvas's settings
 * onto its deps and its frames onto the canvas's outputs:
 *
 *   - a move snaps with the span-guide model the canvas draws (shared
 *     `snapMoveToSpanGuides`: other shapes, then the grid, then ruler guides);
 *   - a resize stays axis-aligned, clamped to {@link MIN_RESIZE}, and snaps its
 *     dragged edges to the grid when snap-to-grid is on;
 *   - a rotate emits whole degrees in [0, 360); Shift snaps it to 15 degrees;
 *   - a lock that arrives mid-gesture ends it where it is.
 *
 * No Angular dependency, so it is unit-testable without mounting the component.
 */
import type {
	GestureController,
	GestureKind,
	InteractionBox,
	PointerLike,
	ResizeHandleId,
	SnapBox,
	SnapGuide,
	SpanSnapRulerGuide,
} from 'ooxml-ui/pptx';
import { createGestureController, snapMoveToSpanGuides } from 'ooxml-ui/pptx';

import { MIN_RESIZE } from './drag-resize';
import type { Box } from './drag-resize';

/** Snap distance (screen pixels) for alignment guides. */
const SNAP_SCREEN_PX = 6;

/** What the canvas supplies to the gestures and receives from them. */
export interface SlideCanvasGestureHost {
	/** On-screen scale (user zoom folded with the auto-fit). */
	scale(): number;
	/** The scaled slide stage, for mapping rotation pointers into slide px. */
	stage(): HTMLElement | undefined;
	/** The element's box when the gesture starts (`rotation` only for a rotate). */
	elementBox(id: string, kind: GestureKind): InteractionBox | undefined;
	/** False once a lock forbids the gesture (re-checked on every move). */
	canInteract(id: string, kind: GestureKind): boolean;
	/** The other boxes a move snaps to. */
	snapOthers(id: string): SnapBox[];
	snapToShape(): boolean;
	snapToGrid(): boolean;
	gridSize(): number;
	/** The ruler guides a move snaps to, or null when guide snapping is off. */
	rulerGuides(): readonly SpanSnapRulerGuide[] | null;
	/** Draw (or clear) the move's span guides. */
	setSnapGuides(guides: readonly SnapGuide[]): void;
	/** The gesture left its dead zone. */
	start(id: string, kind: GestureKind): void;
	/** A move or resize frame. */
	update(id: string, box: Box): void;
	/** A rotate frame, in whole degrees. */
	rotate(id: string, rotation: number): void;
	/** A move or resize finished having moved: reroute what is glued to `ids`. */
	end(ids: readonly string[]): void;
}

export interface SlideCanvasGestures {
	begin(kind: GestureKind, id: string, pointer: PointerLike, handle?: ResizeHandleId): void;
	isActive(): boolean;
	/**
	 * True while the running gesture has left its dead zone. The canvas reads it
	 * from its document `pointerup`, which runs before the gesture's own window
	 * listener ends it.
	 */
	moved(): boolean;
	dispose(): void;
}

export function createSlideCanvasGestures(host: SlideCanvasGestureHost): SlideCanvasGestures {
	let kind: GestureKind = 'move';
	let moved = false;
	const controller: GestureController = createGestureController({
		getScale: () => host.scale() || 1,
		getElementBox: (id, gestureKind) => host.elementBox(id, gestureKind),
		getSiblings: () => [],
		getSnapToGrid: () => host.snapToGrid(),
		getGridSize: () => host.gridSize(),
		getMinSize: () => MIN_RESIZE,
		snapMove: (box, id) => {
			const result = snapMoveToSpanGuides(box, {
				others: host.snapOthers(id),
				threshold: SNAP_SCREEN_PX / (host.scale() || 1),
				snapToShape: host.snapToShape(),
				gridSize: host.snapToGrid() ? host.gridSize() : null,
				rulerGuides: host.rulerGuides(),
			});
			host.setSnapGuides(result.guides);
			return { x: result.x, y: result.y, lines: [] };
		},
		canContinue: (id, gestureKind) => host.canInteract(id, gestureKind),
		getStageOrigin: () => {
			const rect = host.stage()?.getBoundingClientRect();
			return { left: rect?.left ?? 0, top: rect?.top ?? 0 };
		},
		onStart: (id, gestureKind) => {
			moved = true;
			host.start(id, gestureKind);
		},
		onPreview: (transform) => {
			if (kind === 'rotate') {
				host.rotate(transform.id, Math.round(((transform.rotation % 360) + 360) % 360) % 360);
				return;
			}
			const { x, y, width, height } = transform;
			host.update(transform.id, { x, y, width, height });
		},
		onEnd: (_transform, gestureMoved, id) => {
			moved = false;
			host.setSnapGuides([]);
			// Rotation leaves the box (and so every connection site's anchor) alone,
			// so it is not worth a reroute.
			if (gestureMoved && kind !== 'rotate') {
				host.end([id]);
			}
		},
	});
	return {
		begin(gestureKind, id, pointer, handle) {
			if (controller.isActive()) {
				return;
			}
			kind = gestureKind;
			moved = false;
			controller.begin(gestureKind, id, pointer, handle);
		},
		isActive: () => controller.isActive(),
		moved: () => moved,
		dispose: () => controller.dispose(),
	};
}
