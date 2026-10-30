import type { SnapBox, SnapGuide } from '../../geometry/snap-guides';
import type { VisioPage } from '../model';
import { visioSnapTargetBoxes, type VisioSnapOptions } from './guides';
import { visioPageDragDelta, type VisioPagePoint } from './shape-move';
import {
	VISIO_RESIZE_HANDLES,
	visioResizeKeepsProportions,
	visioResizeShape,
	type VisioResizeDragOptions,
	type VisioResizeHandle,
} from './shape-resize';

export interface VisioResizeSnapOptions extends VisioResizeDragOptions {
	/** Dynamic Grid and guides, with the snap distance in page inches. */
	snap: VisioSnapOptions;
}

export interface VisioResizeSnapResult {
	/** The drag end point, moved by the snap so the moving edge sits on its snap line. */
	end: VisioPagePoint;
	/** Alignment hints in the snap frame (y = -page y), as the move hints are. */
	lines: SnapGuide[];
}

type Axis = 'x' | 'y';

const PARALLEL = 1e-9;

/**
 * The page axis a moving edge lies along, or undefined when it is rotated off both axes. An edge
 * whose normal is along page x is vertical, so its position is the x coordinate.
 */
function edgeAxis(normal: { x: number; y: number }): Axis | undefined {
	const length = Math.hypot(normal.x, normal.y);
	if (!(length > 0)) return undefined;
	if (Math.abs(normal.y) / length < PARALLEL) return 'x';
	if (Math.abs(normal.x) / length < PARALLEL) return 'y';
	return undefined;
}

/** The leading edge, centre and trailing edge of a snap box along one axis. */
function lineCandidates(box: SnapBox, axis: Axis): number[] {
	return axis === 'x'
		? [box.x, box.x + box.width / 2, box.x + box.width]
		: [box.y, box.y + box.height / 2, box.y + box.height];
}

/** The box's extent across the snap line, which the hint spans together with the moving point. */
function crossExtent(box: SnapBox, axis: Axis): [number, number] {
	return axis === 'x' ? [box.y, box.y + box.height] : [box.x, box.x + box.width];
}

/**
 * Snap a single-shape resize handle onto other shapes' edges and centres and onto ruler guides, as
 * Dynamic Grid does for moves. The handle's moving edges snap when they lie along a page axis; a
 * rotated edge is left alone, since its snap line is not page-aligned. A proportional drag does not
 * snap, because moving one edge would drag the other off its line. Unchanged input is returned as is.
 */
export function visioResizeSnapEnd(
	page: VisioPage,
	id: string,
	handle: VisioResizeHandle,
	start: VisioPagePoint,
	end: VisioPagePoint,
	options: VisioResizeSnapOptions,
): VisioResizeSnapResult {
	const unchanged: VisioResizeSnapResult = { end, lines: [] };
	const shape = visioResizeShape(page, id);
	const location = VISIO_RESIZE_HANDLES.find((item) => item.id === handle);
	const { snap } = options;
	if (
		!shape ||
		!location ||
		visioResizeKeepsProportions(handle, { ...options, aspect: !!shape?.resizeLocks?.aspect }) ||
		!(snap.shapes || snap.guides) ||
		!(snap.threshold > 0 && Number.isFinite(snap.threshold))
	)
		return unchanged;
	const delta = visioPageDragDelta(start, end);
	if (![delta.x, delta.y].every(Number.isFinite)) return unchanged;
	const [a, b, c, d, e, f] = shape.transform;
	const localX = location.x * shape.width,
		localY = location.y * shape.height;
	// The dragged handle in the snap frame (y = -page y, as the snap boxes are).
	const point = {
		x: a * localX + c * localY + e + delta.x,
		y: -(b * localX + d * localY + f + delta.y),
	};
	const locks = shape.resizeLocks ?? {};
	const moving: { axis: Axis; value: number }[] = [];
	// The x edge runs along local y, so its normal is the local x axis (a, b) in page space.
	const xEdge = location.x !== 0.5 && !locks.width ? edgeAxis({ x: a, y: b }) : undefined;
	// The y edge's normal is the local y axis (c, d).
	const yEdge = location.y !== 0.5 && !locks.height ? edgeAxis({ x: c, y: d }) : undefined;
	for (const axis of [xEdge, yEdge]) if (axis) moving.push({ axis, value: point[axis] });
	const targets = visioSnapTargetBoxes(page, new Set([id]), snap);
	const shift = { x: 0, y: 0 };
	const hits: { axis: Axis; line: number; box: SnapBox }[] = [];
	for (const { axis, value } of moving) {
		let best: { gap: number; line: number; box: SnapBox } | undefined;
		for (const box of targets)
			for (const line of lineCandidates(box, axis)) {
				const gap = line - value;
				if (Math.abs(gap) <= snap.threshold && (!best || Math.abs(gap) < Math.abs(best.gap)))
					best = { gap, line, box };
			}
		if (best) {
			shift[axis] = best.gap;
			hits.push({ axis, line: best.line, box: best.box });
		}
	}
	if (!hits.length) return unchanged;
	// A snap moves the end point by the gap; the snap frame's y runs opposite the drag's y.
	const snapped = { x: point.x + shift.x, y: point.y + shift.y };
	const lines = hits.map(({ axis, line, box }): SnapGuide => {
		const [lo, hi] = crossExtent(box, axis);
		const across = axis === 'x' ? snapped.y : snapped.x;
		return {
			axis,
			pos: line,
			start: Math.min(lo, across),
			end: Math.max(hi, across),
		};
	});
	return {
		end: { x: end.x + shift.x, y: end.y + shift.y },
		lines,
	};
}
