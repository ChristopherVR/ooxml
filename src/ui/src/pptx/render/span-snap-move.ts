import type { SnapBox, SnapGuide } from './snap-guides';
import { computeSnap, snapToGridStep } from './snap-guides';

/**
 * The span-guide move snap: the box snaps to the closest edge or centre of the
 * other boxes (drawing span guides that cover both), then to the grid, then to
 * the user's ruler guides. The Angular canvas draws this model; it moved here
 * from `SlideCanvasComponent` with its move gesture, which now runs on the
 * shared `createGestureController` and plugs this in as its `snapMove`.
 */

/** A user ruler guide: `axis: 'x'` is a vertical line at x = `pos`. */
export interface SpanSnapRulerGuide {
	axis: 'x' | 'y';
	pos: number;
}

export interface SpanSnapMoveOptions {
	/** The other boxes on the stage (the moving one left out). */
	others: readonly SnapBox[];
	/** Snap distance in slide px (the screen distance divided by the zoom). */
	threshold: number;
	snapToShape: boolean;
	/** Grid spacing applied after the shape snap, or null for no grid snap. */
	gridSize: number | null;
	/** Ruler guides applied last, or null when guide snapping is off. */
	rulerGuides: readonly SpanSnapRulerGuide[] | null;
}

/** Where a moving box lands and the span guides to draw for it. */
export interface SpanSnapMoveResult {
	x: number;
	y: number;
	guides: SnapGuide[];
}

/** Snap one axis of a box (`start`, `size`) to the first guide within `threshold`. */
function snapToRulerGuides(
	start: number,
	size: number,
	axis: SpanSnapRulerGuide['axis'],
	guides: readonly SpanSnapRulerGuide[],
	threshold: number,
): number {
	let snapped = start;
	for (const guide of guides) {
		if (guide.axis !== axis) {
			continue;
		}
		for (const candidate of [snapped, snapped + size / 2, snapped + size]) {
			if (Math.abs(candidate - guide.pos) <= threshold) {
				snapped = guide.pos - (candidate - snapped);
				break;
			}
		}
	}
	return snapped;
}

export function snapMoveToSpanGuides(
	box: SnapBox,
	options: SpanSnapMoveOptions,
): SpanSnapMoveResult {
	let { x, y } = box;
	let guides: SnapGuide[] = [];
	if (options.snapToShape) {
		const snap = computeSnap(box, options.others, options.threshold);
		x = snap.x;
		y = snap.y;
		guides = snap.guides;
	}
	// Grid snap: applied after the element-edge snap so the grid takes precedence.
	if (options.gridSize !== null) {
		x = snapToGridStep(x, options.gridSize);
		y = snapToGridStep(y, options.gridSize);
	}
	if (options.rulerGuides) {
		x = snapToRulerGuides(x, box.width, 'x', options.rulerGuides, options.threshold);
		y = snapToRulerGuides(y, box.height, 'y', options.rulerGuides, options.threshold);
	}
	return { x, y, guides };
}
