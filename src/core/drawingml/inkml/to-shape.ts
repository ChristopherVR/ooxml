import type { InkCoordinate } from '../ink-geometry';
import type { RecognizedInkShape } from '../ink-recognition';
import type { DrawingLine } from '../types';

/**
 * Ink to Shape, format-neutral: a recognised drawing (`recognizeInkShape`) as the preset shape
 * Office puts in its place, outlined in the ink's colour and width with no fill. Coordinates are
 * those of the strokes (CSS pixels); each product places the box in its own drawing.
 */
export interface InkShape {
	preset: 'rect' | 'ellipse' | 'triangle' | 'line';
	box: { x: number; y: number; width: number; height: number };
	rotation: number;
	flipH: boolean;
	flipV: boolean;
	line: DrawingLine;
}

const EMU_PER_PIXEL = 9525;

/** The shape that replaces recognised ink drawn in `color` (`#rrggbb`) at `widthPx`. */
export function inkShapeOf(shape: RecognizedInkShape, color: string, widthPx: number): InkShape {
	const line: DrawingLine = {
		widthEmu: Math.max(1, Math.round(widthPx * EMU_PER_PIXEL)),
		fill: {
			kind: 'solid',
			color: { kind: 'srgb', value: color.replace('#', '').toUpperCase(), transforms: [] },
		},
	};
	if ('start' in shape) {
		const { start, end } = shape;
		return {
			preset: 'line',
			box: lineBox(start, end),
			rotation: 0,
			flipH: end.x < start.x,
			flipV: end.y < start.y,
			line: shape.kind === 'arrow' ? { ...line, tailEnd: { type: 'triangle' } } : line,
		};
	}
	const preset = shape.kind === 'rect' || shape.kind === 'ellipse' ? shape.kind : 'triangle';
	return {
		preset,
		...snapped(preset, shape.box, shape.rotation),
		flipH: false,
		flipV: false,
		line,
	};
}

function lineBox(start: InkCoordinate, end: InkCoordinate) {
	return {
		x: Math.min(start.x, end.x),
		y: Math.min(start.y, end.y),
		width: Math.max(1, Math.abs(end.x - start.x)),
		height: Math.max(1, Math.abs(end.y - start.y)),
	};
}

/** Rotations this close to a quarter turn are snapped to it. */
const SNAP_DEGREES = 12;

/**
 * A rectangle or ellipse drawn nearly square to the page is meant square, as Office snaps it: a
 * box within `SNAP_DEGREES` of a quarter turn is turned upright about its centre (its width and
 * height swap for an odd quarter turn). A triangle keeps its rotation, which says where its apex is.
 */
function snapped(
	preset: InkShape['preset'],
	box: InkShape['box'],
	rotation: number,
): Pick<InkShape, 'box' | 'rotation'> {
	const turn = ((rotation % 360) + 360) % 360;
	const quarters = Math.round(turn / 90);
	if (preset === 'triangle' || Math.abs(turn - quarters * 90) > SNAP_DEGREES)
		return { box: { ...box }, rotation };
	if (quarters % 2 === 0) return { box: { ...box }, rotation: 0 };
	const cx = box.x + box.width / 2;
	const cy = box.y + box.height / 2;
	return {
		box: { x: cx - box.height / 2, y: cy - box.width / 2, width: box.height, height: box.width },
		rotation: 0,
	};
}
