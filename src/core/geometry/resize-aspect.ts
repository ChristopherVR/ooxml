import {
	DEFAULT_INTERACTION_MIN_SIZE,
	type BoxTransform,
	type InteractionBox,
	type ResizeHandleId,
} from './element-transform';

const CORNER_HANDLES: ReadonlySet<ResizeHandleId> = new Set(['nw', 'ne', 'se', 'sw']);

/** True for the four corner handles (the ones Shift aspect-locks). */
export function isCornerHandle(handle: ResizeHandleId): boolean {
	return CORNER_HANDLES.has(handle);
}

/**
 * Constrain a corner resize (already computed by the shared `applyResize`) to
 * the start box's aspect ratio: the axis with the larger relative change wins
 * and the other follows, keeping the anchored (opposite) corner fixed. Edge
 * handles are returned unchanged; aspect-locking a one-axis handle would
 * fight the drag direction.
 */
export function lockResizeAspect(
	resized: BoxTransform,
	start: InteractionBox,
	handle: ResizeHandleId,
	minSize: number = DEFAULT_INTERACTION_MIN_SIZE,
): BoxTransform {
	if (!isCornerHandle(handle) || start.width <= 0 || start.height <= 0) {
		return resized;
	}
	const scaleW = resized.width / start.width;
	const scaleH = resized.height / start.height;
	const scale = Math.abs(scaleW - 1) >= Math.abs(scaleH - 1) ? scaleW : scaleH;

	let width = start.width * scale;
	let height = start.height * scale;
	// Clamp preserving the ratio: bump the scale so both axes stay >= minSize.
	if (width < minSize || height < minSize) {
		const clampScale = Math.max(minSize / start.width, minSize / start.height);
		width = start.width * clampScale;
		height = start.height * clampScale;
	}

	const affectsLeft = handle === 'nw' || handle === 'sw';
	const affectsTop = handle === 'nw' || handle === 'ne';
	return {
		x: affectsLeft ? start.x + start.width - width : start.x,
		y: affectsTop ? start.y + start.height - height : start.y,
		width,
		height,
		rotation: resized.rotation,
	};
}
