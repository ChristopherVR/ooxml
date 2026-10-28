/** Format-neutral move and resize geometry, extracted from the PowerPoint editing overlay. */
export const DEFAULT_INTERACTION_MIN_SIZE = 12;

/** Axis-aligned box in element (unscaled) coordinate space, plus rotation. */
export interface InteractionBox {
	x: number;
	y: number;
	width: number;
	height: number;
	/** Degrees. Defaults to 0 when omitted. */
	rotation?: number;
}

/** Result of a drag/resize/rotate computation. */
export interface BoxTransform {
	x: number;
	y: number;
	width: number;
	height: number;
	rotation: number;
}

/** The eight resize handles, named by compass direction. */
export type ResizeHandleId = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';

export const RESIZE_HANDLES: readonly ResizeHandleId[] = [
	'nw',
	'n',
	'ne',
	'e',
	'se',
	's',
	'sw',
	'w',
] as const;

/** Where one resize handle sits on the box, and the cursor it shows. */
export interface ResizeHandleGeometry {
	/** CSS cursor value (not a class name). */
	cursor: string;
	/** Fractional x within the box: 0 = left edge, 1 = right edge. */
	fx: number;
	/** Fractional y within the box: 0 = top edge, 1 = bottom edge. */
	fy: number;
}

/**
 * The canonical placement + cursor for all eight handles.
 *
 * All five bindings had their own copy of this table: Vue as `HANDLE_META`,
 * Svelte and Vanilla as a `HANDLE_CURSORS` / `HANDLE_POSITIONS` pair, Angular
 * as a `switch` plus a string-matching `handleAnchor`, and React as `cursor-*`
 * Tailwind classes on hand-placed corner/edge lists. Five copies of eight
 * constants is five chances for one binding to put a handle on the wrong corner
 * or show a resize cursor pointing the wrong way, with nothing to catch it.
 */
export const RESIZE_HANDLE_GEOMETRY: Readonly<Record<ResizeHandleId, ResizeHandleGeometry>> = {
	nw: { cursor: 'nwse-resize', fx: 0, fy: 0 },
	n: { cursor: 'ns-resize', fx: 0.5, fy: 0 },
	ne: { cursor: 'nesw-resize', fx: 1, fy: 0 },
	e: { cursor: 'ew-resize', fx: 1, fy: 0.5 },
	se: { cursor: 'nwse-resize', fx: 1, fy: 1 },
	s: { cursor: 'ns-resize', fx: 0.5, fy: 1 },
	sw: { cursor: 'nesw-resize', fx: 0, fy: 1 },
	w: { cursor: 'ew-resize', fx: 0, fy: 0.5 },
};

/**
 * Rotate-handle stem length in SCREEN px, i.e. constant at any zoom.
 *
 * The overlay lives inside the zoom-scaled stage, so this is divided by the
 * zoom wherever it is applied; a stem that scaled with the zoom would drift
 * away from the box as the user zoomed out.
 */
export const ROTATE_STEM_PX = 24;

export interface Point {
	x: number;
	y: number;
}

export interface ResizeOptions {
	/** Minimum width/height in element px. Defaults to `DEFAULT_INTERACTION_MIN_SIZE`. */
	minSize?: number;
	/** Per-axis minimums, in the same units as the box; override minSize when supplied. */
	minWidth?: number;
	minHeight?: number;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const DEG_TO_RAD = Math.PI / 180;

function rotationOf(box: InteractionBox): number {
	return box.rotation ?? 0;
}

/**
 * Rotate a screen-space delta into element-local space.
 *
 * Given a vector expressed in the canvas (screen) frame and the element's
 * rotation in degrees, return the same vector expressed in the element's local
 * (un-rotated) frame. This is the inverse rotation of `rotationDeg`.
 */
export function rotateDelta(dx: number, dy: number, rotationDeg: number): Point {
	if (!rotationDeg) {
		return { x: dx, y: dy };
	}
	const rad = -rotationDeg * DEG_TO_RAD;
	const cos = Math.cos(rad);
	const sin = Math.sin(rad);
	return {
		x: dx * cos - dy * sin,
		y: dx * sin + dy * cos,
	};
}

/** Center point of a box in element coordinates. */
export function boxCenter(box: InteractionBox): Point {
	return { x: box.x + box.width / 2, y: box.y + box.height / 2 };
}

// ---------------------------------------------------------------------------
// Drag (move)
// ---------------------------------------------------------------------------

/**
 * Apply a screen-space pointer delta as a plain move of the whole box.
 *
 * Position is stored axis-aligned, so rotation does NOT affect a move  -  we only
 * scale the screen delta down by `zoom` to get element px.
 */
export function applyDragDelta(
	box: InteractionBox,
	dxScreen: number,
	dyScreen: number,
	zoom: number,
): BoxTransform {
	const scale = zoom || 1;
	return {
		x: box.x + dxScreen / scale,
		y: box.y + dyScreen / scale,
		width: box.width,
		height: box.height,
		rotation: rotationOf(box),
	};
}

// ---------------------------------------------------------------------------
// Resize
// ---------------------------------------------------------------------------

/**
 * Compute the new geometry for a resize drag on a given handle.
 *
 * `dxScreen`/`dyScreen` are the *total* pointer delta from the gesture start,
 * in screen px. They are scaled by `zoom` into element px and (when the element
 * is rotated) rotated into the element's local frame so the handle behaves
 * intuitively. The min-size clamp uses `opts.minSize` (default
 * `DEFAULT_INTERACTION_MIN_SIZE`).
 *
 * The returned box keeps the element's `rotation` unchanged. x/y describe the
 * new axis-aligned top-left in element space; for rotated elements this matches
 * the React behaviour of resizing in local space under a CSS transform.
 */
export function applyResize(
	box: InteractionBox,
	handle: ResizeHandleId,
	dxScreen: number,
	dyScreen: number,
	zoom: number,
	opts: ResizeOptions = {},
): BoxTransform {
	const scale = zoom || 1;
	const minSize = opts.minSize ?? DEFAULT_INTERACTION_MIN_SIZE;
	const minWidth = opts.minWidth ?? minSize;
	const minHeight = opts.minHeight ?? minSize;
	const rotation = rotationOf(box);

	// Screen delta -> element px -> element-local frame.
	const local = rotateDelta(dxScreen / scale, dyScreen / scale, rotation);
	const dx = local.x;
	const dy = local.y;

	let newX = box.x;
	let newY = box.y;
	let newW = box.width;
	let newH = box.height;

	const affectsLeft = handle === 'nw' || handle === 'w' || handle === 'sw';
	const affectsRight = handle === 'ne' || handle === 'e' || handle === 'se';
	const affectsTop = handle === 'nw' || handle === 'n' || handle === 'ne';
	const affectsBottom = handle === 'sw' || handle === 's' || handle === 'se';

	if (affectsRight) {
		newW = Math.max(minWidth, box.width + dx);
	} else if (affectsLeft) {
		newW = Math.max(minWidth, box.width - dx);
		newX = box.x + (box.width - newW);
	}

	if (affectsBottom) {
		newH = Math.max(minHeight, box.height + dy);
	} else if (affectsTop) {
		newH = Math.max(minHeight, box.height - dy);
		newY = box.y + (box.height - newH);
	}

	return { x: newX, y: newY, width: newW, height: newH, rotation };
}
