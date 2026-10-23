import { lockResizeAspect } from './editor-geometry';
import type { InteractionBox, Point, ResizeHandleId } from './element-interaction';
import {
	applyDragDelta,
	applyResize,
	boxCenter,
	snapAngle,
	snapBoxToGrid,
} from './element-interaction';
import { createRotationDrag } from './rotation-drag';
import type { SnapGuideInput, SnapLine, SnapSibling } from './snap-guides';
import { computeSnapToShape, snapToGridStep } from './snap-guides';

/**
 * The pure per-frame maths behind `createGestureController`: one function per
 * gesture kind that turns the gesture's start box and the pointer's travel
 * into the next box. Split from `editor-gestures.ts`, which owns only the
 * pointer lifecycle, so each rule is testable without dispatching events.
 */

/** Live geometry emitted during and at the end of a gesture (element px). */
export interface GestureTransform {
	id: string;
	x: number;
	y: number;
	width: number;
	height: number;
	rotation: number;
}

/** A snapped move position and the alignment lines to draw for it. */
export interface GestureSnap {
	x: number;
	y: number;
	lines: readonly SnapLine[];
}

/** What a move frame needs to know about the stage's snap settings. */
export interface MoveFrameOptions {
	scale: number;
	snapToGrid: boolean;
	gridSize: number;
	snapToShape: boolean;
	siblings: readonly SnapSibling[];
	/** Ids the shape snap ignores: the dragged element, or a whole multi-selection. */
	excludeIds: ReadonlySet<string>;
	guides: readonly SnapGuideInput[];
	/**
	 * A binding's own snap model, replacing the grid and shape snap above (the
	 * Angular canvas keeps its span guides and ruler-guide snap this way).
	 */
	snapMove?: (box: InteractionBox, id: string) => GestureSnap;
}

export function toGestureTransform(id: string, box: InteractionBox): GestureTransform {
	return {
		id,
		x: box.x,
		y: box.y,
		width: box.width,
		height: box.height,
		rotation: box.rotation ?? 0,
	};
}

/** The box a move of `dx`/`dy` screen px lands on, snapped, plus its guide lines. */
export function computeMoveFrame(
	id: string,
	start: InteractionBox,
	dx: number,
	dy: number,
	options: MoveFrameOptions,
): [GestureTransform, readonly SnapLine[]] {
	let box: InteractionBox = applyDragDelta(start, dx, dy, options.scale);
	if (options.snapMove) {
		const snapped = options.snapMove(box, id);
		return [{ ...toGestureTransform(id, box), x: snapped.x, y: snapped.y }, snapped.lines];
	}
	if (options.snapToGrid) {
		box = {
			...box,
			x: snapToGridStep(box.x, options.gridSize),
			y: snapToGridStep(box.y, options.gridSize),
		};
	}
	// Snap to sibling edges/centres (skipped for rotated boxes, whose AABB
	// no longer matches the visual outline).
	if (!box.rotation && options.snapToShape) {
		const snapped = computeSnapToShape(
			box.x,
			box.y,
			box.width,
			box.height,
			options.siblings,
			new Set(options.excludeIds),
			options.guides,
		);
		return [{ ...toGestureTransform(id, box), x: snapped.x, y: snapped.y }, snapped.lines];
	}
	return [toGestureTransform(id, box), []];
}

/** What a resize frame needs beyond the handle and the travel. */
export interface ResizeFrameOptions {
	scale: number;
	/** Shift held: a corner handle keeps the start box's aspect ratio. */
	lockAspect: boolean;
	/** Smallest width/height the resize may reach; the geometry default when absent. */
	minSize?: number;
	/** Grid spacing the dragged edges snap to, or null for no grid snap. */
	gridSize: number | null;
}

/** The box a resize of `dx`/`dy` screen px on `handle` lands on. */
export function computeResizeFrame(
	id: string,
	start: InteractionBox,
	handle: ResizeHandleId | undefined,
	dx: number,
	dy: number,
	options: ResizeFrameOptions,
): GestureTransform {
	if (!handle) {
		return toGestureTransform(id, start);
	}
	const { minSize } = options;
	let box = applyResize(
		start,
		handle,
		dx,
		dy,
		options.scale,
		minSize === undefined ? {} : { minSize },
	);
	if (options.lockAspect) {
		box = lockResizeAspect(box, start, handle, minSize);
	}
	if (options.gridSize !== null) {
		// Per edge: the anchored edges stay put, the dragged ones land on the grid.
		box = { ...box, ...snapBoxToGrid(box, handle, options.gridSize, minSize) };
	}
	return toGestureTransform(id, box);
}

/** A rotation drag anchored to the press (see `createRotationDrag`). */
export function startRotationFrame(
	start: InteractionBox,
	press: Point,
): (pointer: Point) => number {
	return createRotationDrag(boxCenter(start), press, start.rotation ?? 0);
}

/** The box with the rotation the pointer asks for; Shift snaps to 15 degrees. */
export function computeRotateFrame(
	id: string,
	start: InteractionBox,
	rotate: ((pointer: Point) => number) | undefined,
	pointer: Point,
	shift: boolean,
): GestureTransform {
	let angle = rotate?.(pointer) ?? start.rotation ?? 0;
	if (shift) {
		angle = snapAngle(angle);
	}
	return { ...toGestureTransform(id, start), rotation: angle };
}
