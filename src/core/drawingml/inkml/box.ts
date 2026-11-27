/**
 * Where an ink drawing sits and how its own coordinates map onto that box, as Word and Excel
 * both use it: whether a new stroke joins the drawing, whether the eraser reaches it, a point in
 * the drawing's coordinates, and the box fitted to the strokes it keeps.
 *
 * @module drawingml/inkml/box
 */
import type { InkCoordinate } from '../ink-geometry';
import { inkStrokesBounds, transformInkStrokes, type InkBox } from './edit';
import type { InkStroke } from './types';

/** A stroke drawn this soon (ms) and this close (px) to a drawing joins it, as Office collects ink. */
export const INK_JOIN_MS = 2000;
export const INK_JOIN_PX = 48;

/** The size of a drawing's own coordinate space. */
export interface InkSpace {
	width: number;
	height: number;
}

/** Whether a stroke with `bounds`, drawn `elapsed` ms after the last, joins the drawing at `box`. */
export function inkJoins(box: InkBox, bounds: InkBox, elapsed: number): boolean {
	return (
		elapsed < INK_JOIN_MS &&
		bounds.x < box.x + box.width + INK_JOIN_PX &&
		bounds.x + bounds.width > box.x - INK_JOIN_PX &&
		bounds.y < box.y + box.height + INK_JOIN_PX &&
		bounds.y + bounds.height > box.y - INK_JOIN_PX
	);
}

/** Whether a point is within `reach` of a drawing's box. */
export function inkBoxReaches(box: InkBox, point: InkCoordinate, reach: number): boolean {
	return (
		point.x >= box.x - reach &&
		point.x <= box.x + box.width + reach &&
		point.y >= box.y - reach &&
		point.y <= box.y + box.height + reach
	);
}

/** A point (in the box's coordinates) in the drawing's own, and the factor for lengths. */
export function inkLocalPoint(
	box: InkBox,
	space: InkSpace,
	point: InkCoordinate,
): { point: InkCoordinate; factor: number } {
	const factor = space.width / (box.width || 1);
	return {
		point: {
			x: (point.x - box.x) * factor,
			y: (point.y - box.y) * (space.height / (box.height || 1)),
		},
		factor,
	};
}

/** `box` on whole pixels, covering it. */
export function snapInkBox(box: InkBox): InkBox {
	const x = Math.floor(box.x);
	const y = Math.floor(box.y);
	return {
		x,
		y,
		width: Math.max(1, Math.ceil(box.x + box.width - x)),
		height: Math.max(1, Math.ceil(box.y + box.height - y)),
	};
}

/**
 * The drawing at `box` (its own coordinates `space`) keeping only `strokes`: its box fitted to
 * them, as a new stroke fits a growing drawing, and the strokes in the fitted box's space.
 * `snap` puts the box on whole pixels. Undefined when no stroke is left.
 */
export function fitInkBox(
	strokes: readonly InkStroke[],
	space: InkSpace,
	box: InkBox,
	snap = false,
): { box: InkBox; strokes: InkStroke[]; width: number; height: number } | undefined {
	const bounds = inkStrokesBounds(strokes);
	if (!bounds) return undefined;
	const sx = box.width / (space.width || 1);
	const sy = box.height / (space.height || 1);
	const scaled = {
		x: bounds.x * sx,
		y: bounds.y * sy,
		width: bounds.width * sx,
		height: bounds.height * sy,
	};
	const inner = snap ? snapInkBox(scaled) : scaled;
	return {
		box: { x: box.x + inner.x, y: box.y + inner.y, width: inner.width, height: inner.height },
		strokes: transformInkStrokes(strokes, -inner.x / sx, -inner.y / sy),
		width: inner.width / sx,
		height: inner.height / sy,
	};
}
