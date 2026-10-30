import { computeSnap, type SnapBox, type SnapGuide } from '../../geometry/snap-guides';
import type { VisioEdit } from '../edit-commands';
import type { VisioPage } from '../model';
import { visioShapeSnapBox, visioSnapTargets } from './guides';
import type { VisioPagePoint } from './shape-move';
import {
	VISIO_RESIZE_HANDLES,
	type VisioResizeDragOptions,
	type VisioResizeFrame,
	type VisioResizeHandle,
} from './shape-resize';
import { visioResizeSelectionDrag } from './selection-resize';

/** Dynamic Grid and Guides while a handle is dragged, as for a move (see `visioSnapMoveDelta`). */
export interface VisioResizeSnapOptions {
	/** Dynamic Grid: the moving edge snaps to other shapes' edges and centres, and size hints show. */
	shapes: boolean;
	/** Shown guides: the moving edge snaps to them. */
	guides: boolean;
	/** Page inches; a snap happens within this distance. */
	threshold: number;
}

/** A nearby shape whose width or height equals the frame's; `box` is drawn as the hint. */
export interface VisioResizeSizeMatch {
	axis: 'x' | 'y';
	shapeId: string;
	box: SnapBox;
}

export interface VisioSnapResizeResult {
	frame: VisioResizeFrame;
	commands: VisioEdit[];
	members: { id: string; frame: VisioResizeFrame }[];
	/** Alignment lines in the snap box convention (y = -page y), for the same overlay a move uses. */
	lines: SnapGuide[];
	sizeMatches: VisioResizeSizeMatch[];
}

/** A frame's nearby shapes are those within this many page inches of its box. */
const NEARBY = 1;
/** Two lengths count as equal within this many page inches (about 0.03 mm). */
const SAME_LENGTH = 1e-3;

/** The point a handle sits on in the snap box convention (y = -page y). */
export function visioResizeHandlePoint(
	frame: Pick<VisioResizeFrame, 'transform' | 'width' | 'height'>,
	location: { x: number; y: number },
): VisioPagePoint {
	const [a, b, c, d, e, f] = frame.transform;
	const lx = location.x * frame.width,
		ly = location.y * frame.height;
	return { x: a * lx + c * ly + e, y: -(b * lx + d * ly + f) };
}

/**
 * Shapes near a resized frame whose width or height equals the frame's. Visio's size-match hint
 * shows them while Dynamic Grid is on. The moving shapes are left out.
 */
export function visioResizeSizeMatches(
	page: VisioPage,
	ids: readonly string[],
	frame: VisioResizeFrame,
): VisioResizeSizeMatch[] {
	const box = visioShapeSnapBox(frame);
	if (!box) return [];
	const moving = new Set(ids);
	const matches: VisioResizeSizeMatch[] = [];
	for (const shape of page.shapes) {
		if (moving.has(shape.id) || shape.hidden || shape.visibility?.guide) continue;
		const other = visioShapeSnapBox(shape);
		if (!other || !overlapsWithin(box, other, NEARBY)) continue;
		if (Math.abs(shape.width - frame.width) <= SAME_LENGTH)
			matches.push({ axis: 'x', shapeId: shape.id, box: other });
		if (Math.abs(shape.height - frame.height) <= SAME_LENGTH)
			matches.push({ axis: 'y', shapeId: shape.id, box: other });
	}
	return matches;
}

function overlapsWithin(a: SnapBox, b: SnapBox, margin: number): boolean {
	return (
		a.x - margin <= b.x + b.width &&
		b.x <= a.x + a.width + margin &&
		a.y - margin <= b.y + b.height &&
		b.y <= a.y + a.height + margin
	);
}

/**
 * A resize drag with Dynamic Grid and Guides. The handle's moving edge snaps to the nearest edge
 * or centre within the threshold, on each axis the handle moves; the drag is then re-solved with
 * the pointer moved by that snap, so the edge lands on the line. Lines report what was aligned.
 * Without `snap` (Alt held, or snapping off) the result is the plain drag.
 */
export function visioSnapResizeDrag(
	page: VisioPage,
	ids: readonly string[],
	handle: VisioResizeHandle,
	start: VisioPagePoint,
	end: VisioPagePoint,
	options: VisioResizeDragOptions & { snap?: VisioResizeSnapOptions } = {},
): VisioSnapResizeResult | undefined {
	const { snap, ...drag } = options;
	const first = visioResizeSelectionDrag(page, ids, handle, start, end, drag);
	if (!first) return undefined;
	const plain = { ...first, lines: [], sizeMatches: [] };
	if (!snap || !(snap.threshold > 0) || !(snap.shapes || snap.guides)) return plain;
	const location = VISIO_RESIZE_HANDLES.find((item) => item.id === handle);
	if (!location) return undefined;
	const useX = location.x !== 0.5,
		useY = location.y !== 0.5;
	const point = visioResizeHandlePoint(first.frame, location);
	const targets: SnapBox[] = visioSnapTargets(page, new Set(ids), snap);
	const hit = computeSnap({ ...point, width: 0, height: 0 }, targets, snap.threshold);
	const lines = hit.guides.filter((line) => (line.axis === 'x' ? useX : useY));
	const dx = useX ? hit.x - point.x : 0,
		dy = useY ? hit.y - point.y : 0;
	let result = first;
	if (dx || dy) {
		// The snap box and the pointer share a direction on both axes, so the pointer moves by the snap.
		const moved = visioResizeSelectionDrag(
			page,
			ids,
			handle,
			start,
			{ x: end.x + dx, y: end.y + dy },
			drag,
		);
		// A snap the re-solved drag cannot honour is dropped, with its lines.
		if (!moved) return plain;
		result = moved;
	}
	return {
		...result,
		lines,
		sizeMatches: snap.shapes ? visioResizeSizeMatches(page, ids, result.frame) : [],
	};
}
