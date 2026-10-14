import { MIN_ELEMENT_SIZE } from 'ooxml-core/pptx';
import type { ResizeHandleId } from 'ooxml-core/geometry';
export type {
	InteractionBox,
	BoxTransform,
	ResizeHandleId,
	ResizeHandleGeometry,
	Point,
	ResizeOptions,
} from 'ooxml-core/geometry';
export {
	RESIZE_HANDLES,
	RESIZE_HANDLE_GEOMETRY,
	ROTATE_STEM_PX,
	rotateDelta,
	boxCenter,
	applyDragDelta,
	applyResize,
} from 'ooxml-core/geometry';

// ---------------------------------------------------------------------------
// Rotation
// ---------------------------------------------------------------------------

/**
 * Compute the rotation angle (degrees) for a pointer relative to a center.
 *
 * 0deg points straight up (the rotate handle sits above the element), matching
 * the screen "up = -Y" convention. The result is normalised to [0, 360).
 */
export { computeRotation } from 'ooxml-core/geometry';

/** Shift-drag rotation snapping; the implementation lives in ooxml-core/geometry. */
export { snapAngle } from 'ooxml-core/geometry';

// ---------------------------------------------------------------------------
// Grid snapping (resize)
// ---------------------------------------------------------------------------

/** Bounding box for grid-snap, in element coordinates. */
export interface GridSnapBox {
	x: number;
	y: number;
	width: number;
	height: number;
}

/**
 * Snap the edges of a resize box to a pixel grid.
 *
 * Each of the four edges touched by `handle` is rounded to the nearest multiple
 * of `gridSpacingPx`; the opposite (anchored) edge is left untouched, and the
 * dimension is recomputed from the two edges (clamped to `minSize`). This mirrors
 * the per-edge snap React applies after computing a raw resize: a handle that
 * affects the right/bottom edge snaps that edge and keeps x/y fixed, while a
 * handle that affects the left/top edge snaps the origin and grows/shrinks the
 * size to compensate.
 *
 * `handle === null` (a plain move, no resize) returns the box unchanged.
 */
export function snapBoxToGrid(
	box: GridSnapBox,
	handle: ResizeHandleId | null,
	gridSpacingPx: number,
	minSize = MIN_ELEMENT_SIZE,
): GridSnapBox {
	if (handle === null || gridSpacingPx <= 0) {
		return box;
	}

	let { x, y, width, height } = box;
	const gs = gridSpacingPx;

	const affectsLeft = handle === 'nw' || handle === 'w' || handle === 'sw';
	const affectsRight = handle === 'ne' || handle === 'e' || handle === 'se';
	const affectsTop = handle === 'nw' || handle === 'n' || handle === 'ne';
	const affectsBottom = handle === 'sw' || handle === 's' || handle === 'se';

	// Snap right edge (x fixed).
	if (affectsRight) {
		const right = Math.round((x + width) / gs) * gs;
		width = Math.max(minSize, right - x);
	}
	// Snap left edge (origin moves, width compensates).
	if (affectsLeft) {
		const snappedX = Math.round(x / gs) * gs;
		width = Math.max(minSize, width + (x - snappedX));
		x = snappedX;
	}
	// Snap bottom edge (y fixed).
	if (affectsBottom) {
		const bottom = Math.round((y + height) / gs) * gs;
		height = Math.max(minSize, bottom - y);
	}
	// Snap top edge (origin moves, height compensates).
	if (affectsTop) {
		const snappedY = Math.round(y / gs) * gs;
		height = Math.max(minSize, height + (y - snappedY));
		y = snappedY;
	}

	return { x, y, width, height };
}

// ---------------------------------------------------------------------------
// Marquee (rubber-band) selection
// ---------------------------------------------------------------------------

/** A marquee drag described by its start and current corner (any order). */
export interface MarqueeRect {
	startX: number;
	startY: number;
	currentX: number;
	currentY: number;
}

/** An element reduced to its id + bounding box for marquee hit-testing. */
export interface MarqueeElementRect {
	id: string;
	x: number;
	y: number;
	width: number;
	height: number;
}

/**
 * Compute which element ids are hit by a marquee selection rectangle.
 *
 * The marquee corners are normalised (start/current may be in any order). An
 * element is hit when its box (its width/height clamped to at least `minSize`,
 * so zero-size shapes are still selectable) overlaps the marquee AABB. A marquee
 * smaller than 3px in *both* dimensions is treated as a click, returning `[]`.
 */
export function computeMarqueeHitIds(
	marquee: MarqueeRect,
	elements: readonly MarqueeElementRect[],
	minSize = MIN_ELEMENT_SIZE,
): string[] {
	const minX = Math.min(marquee.startX, marquee.currentX);
	const minY = Math.min(marquee.startY, marquee.currentY);
	const maxX = Math.max(marquee.startX, marquee.currentX);
	const maxY = Math.max(marquee.startY, marquee.currentY);
	const w = maxX - minX;
	const h = maxY - minY;
	if (w <= 3 && h <= 3) {
		return [];
	}
	return elements
		.filter((el) => {
			const eMinX = el.x;
			const eMinY = el.y;
			const eMaxX = el.x + Math.max(el.width, minSize);
			const eMaxY = el.y + Math.max(el.height, minSize);
			return !(eMaxX < minX || eMinX > maxX || eMaxY < minY || eMinY > maxY);
		})
		.map((el) => el.id);
}

/**
 * Merge an additive (shift-drag) marquee result into an existing selection,
 * de-duplicating ids while preserving first-seen order (base ids first).
 */
export function mergeAdditiveSelection(
	baseSelectionIds: readonly string[] | undefined,
	hitIds: readonly string[],
): string[] {
	return Array.from(new Set([...(baseSelectionIds ?? []), ...hitIds]));
}

/**
 * True when a press should EXTEND the selection rather than replace it.
 *
 * PowerPoint accepts Shift and Ctrl (Cmd on macOS) for this, and the bindings
 * disagreed about which: Vanilla honoured Shift alone, so Ctrl+click on Windows
 * silently destroyed a multi-selection the user was half-way through building.
 * One predicate, so a modifier can never mean "extend" in one viewer and
 * "start over" in another.
 */
export function isAdditiveSelectionPress(event: {
	shiftKey: boolean;
	ctrlKey?: boolean;
	metaKey?: boolean;
}): boolean {
	return event.shiftKey || event.ctrlKey === true || event.metaKey === true;
}
