// Drawing anchor geometry: converts the anchors of pictures, charts and shapes (cell markers with
// EMU offsets, or a one-cell marker with an EMU extent) to plane pixel rectangles at the metrics'
// zoom, and back. Ported from xlsx-viewer `packages/web-component/src/grid/drawings.ts`
// (`anchorBox`, `boxAnchor`) and the extent maths of `commands/insert.ts`.
import { EMU_PER_PIXEL } from '../../units/constants.js';
import type { DrawingAnchor, Worksheet } from '../model.js';
import { createGridMetrics, type GridMetrics } from './metrics.js';

/** A rectangle on the grid plane in pixels at the metrics' zoom (origin: top-left of A1). */
export interface PixelBox {
	x: number;
	y: number;
	w: number;
	h: number;
}

/** Which anchor form `pixelBoxToAnchor` produces. */
export type AnchorKind = 'twoCell' | 'oneCell';

/** Extent used when a one-cell anchor has no `ext` (200 px, Excel's default chart size). */
export const DEFAULT_ANCHOR_EXTENT_EMU = 1905000;

/** Smallest edge of a two-cell box, in pixels, so a collapsed anchor stays selectable. */
export const MIN_TWO_CELL_PIXELS = 4;

type Marker = DrawingAnchor['from'];

/**
 * The anchor form of an existing anchor. Absolute anchors are read into the model as a one-cell
 * marker on A1 with their position as offsets, so they report `oneCell`.
 */
export function anchorKind(anchor: DrawingAnchor): AnchorKind {
	return anchor.to ? 'twoCell' : 'oneCell';
}

const zoomFactor = (metrics: GridMetrics): number => metrics.zoom / 100;

function markerPoint(metrics: GridMetrics, marker: Marker, f: number): { x: number; y: number } {
	return {
		x: metrics.colLeft(marker.col) + (marker.colOffset / EMU_PER_PIXEL) * f,
		y: metrics.rowTop(marker.row) + (marker.rowOffset / EMU_PER_PIXEL) * f,
	};
}

/**
 * The plane rectangle of an anchor. Two-cell anchors span their markers (at least
 * `MIN_TWO_CELL_PIXELS` per edge); one-cell and absolute anchors start at their marker and take
 * their extent (`DEFAULT_ANCHOR_EXTENT_EMU` when missing). `metrics` defaults to the sheet's grid
 * at its own zoom.
 */
export function anchorToPixelBox(
	sheet: Worksheet,
	anchor: DrawingAnchor,
	metrics: GridMetrics = createGridMetrics(sheet),
): PixelBox {
	const f = zoomFactor(metrics);
	const from = markerPoint(metrics, anchor.from, f);
	if (anchor.to) {
		const to = markerPoint(metrics, anchor.to, f);
		return {
			x: from.x,
			y: from.y,
			w: Math.max(MIN_TWO_CELL_PIXELS, to.x - from.x),
			h: Math.max(MIN_TWO_CELL_PIXELS, to.y - from.y),
		};
	}
	return {
		x: from.x,
		y: from.y,
		w: ((anchor.ext?.cx ?? DEFAULT_ANCHOR_EXTENT_EMU) / EMU_PER_PIXEL) * f,
		h: ((anchor.ext?.cy ?? DEFAULT_ANCHOR_EXTENT_EMU) / EMU_PER_PIXEL) * f,
	};
}

function pointMarker(metrics: GridMetrics, px: number, py: number, f: number): Marker {
	const col = metrics.colAt(Math.max(0, px));
	const row = metrics.rowAt(Math.max(0, py));
	return {
		col,
		row,
		colOffset: Math.round(((px - metrics.colLeft(col)) / f) * EMU_PER_PIXEL),
		rowOffset: Math.round(((py - metrics.rowTop(row)) / f) * EMU_PER_PIXEL),
	};
}

/**
 * The anchor that puts an object at a plane rectangle. `like` keeps the form of an existing
 * anchor (pass the object's current anchor when moving or resizing it) or names one; the default
 * is a one-cell anchor. Offsets and extents are whole EMU at 100% zoom.
 */
export function pixelBoxToAnchor(
	sheet: Worksheet,
	box: PixelBox,
	metrics: GridMetrics = createGridMetrics(sheet),
	like: DrawingAnchor | AnchorKind = 'oneCell',
): DrawingAnchor {
	const f = zoomFactor(metrics);
	const kind = typeof like === 'string' ? like : anchorKind(like);
	const next: DrawingAnchor = { from: pointMarker(metrics, box.x, box.y, f) };
	if (kind === 'twoCell') next.to = pointMarker(metrics, box.x + box.w, box.y + box.h, f);
	else next.ext = pixelSizeToExtent(box.w / f, box.h / f);
	return next;
}

/** An EMU extent for a size in pixels at 100% zoom (rounded to whole EMU). */
export function pixelSizeToExtent(width: number, height: number): { cx: number; cy: number } {
	return { cx: Math.round(width * EMU_PER_PIXEL), cy: Math.round(height * EMU_PER_PIXEL) };
}

/**
 * A one-cell anchor at a cell for a picture of `width` x `height` pixels, scaled down to at most
 * `maxWidth` pixels wide (default 640, as Insert > Pictures does) keeping its aspect ratio.
 */
export function pictureAnchorAt(
	cell: { row: number; col: number },
	width: number,
	height: number,
	maxWidth = 640,
): DrawingAnchor {
	const scale = Math.min(1, maxWidth / Math.max(1, width));
	return {
		from: { row: cell.row, col: cell.col, rowOffset: 0, colOffset: 0 },
		ext: pixelSizeToExtent(width * scale, height * scale),
	};
}
