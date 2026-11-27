// The box of a legacy note (the VML shape Excel shows a cell comment in): where it sits and how
// large it is. Excel opens a new note one column right of its cell and a row above, 144 by 79
// pixels; the box here keeps that place and width and grows down to hold the note's text, so a
// long note is not clipped to its first lines.
import type { GridMetrics } from '../layout/metrics';

/** Excel's new-note box, in pixels at 100% zoom: 108 by 59.25 points. */
export const NOTE_WIDTH_PX = 144;
export const NOTE_MIN_HEIGHT_PX = 79;
/** The note's offset from the column right of its cell and from the row above it, in pixels. */
const LEFT_OFFSET_PX = 15;
const TOP_OFFSET_PX = 10;
/**
 * An estimate of the default note text (Tahoma 9 pt): its average character advance, its line
 * height and the box's text insets, in pixels. Text is not measured, so a box is approximate.
 */
const CHAR_PX = 6;
const LINE_PX = 15;
const INSET_X_PX = 14;
const INSET_Y_PX = 8;

export interface NoteBox {
	/** `x:Anchor`: left column, offset, top row, offset, right column, offset, bottom row, offset. */
	anchor: string;
	/** The shape's CSS position and size, in points. */
	style: string;
}

/** The lines the note's text wraps to in a box of the default width. */
export function noteLines(text: string): number {
	const perLine = Math.max(1, Math.floor((NOTE_WIDTH_PX - INSET_X_PX) / CHAR_PX));
	return text
		.split('\n')
		.reduce((sum, line) => sum + Math.max(1, Math.ceil(line.length / perLine)), 0);
}

const points = (px: number) => `${Math.round(px * 75) / 100}pt`;

/** The box of the note on the cell at `row`, `col` holding `text`, on the sheet's grid. */
export function noteBox(grid: GridMetrics, row: number, col: number, text: string): NoteBox {
	const height = Math.max(NOTE_MIN_HEIGHT_PX, noteLines(text) * LINE_PX + INSET_Y_PX);
	const left = col + 1;
	const top = Math.max(0, row - 1);
	const x = grid.colLeft(left) + LEFT_OFFSET_PX;
	const y = grid.rowTop(top) + TOP_OFFSET_PX;
	const right = grid.colAt(x + NOTE_WIDTH_PX);
	const bottom = grid.rowAt(y + height);
	const offset = (value: number) => Math.max(0, Math.round(value));
	return {
		anchor: [
			left,
			LEFT_OFFSET_PX,
			top,
			TOP_OFFSET_PX,
			right,
			offset(x + NOTE_WIDTH_PX - grid.colLeft(right)),
			bottom,
			offset(y + height - grid.rowTop(bottom)),
		].join(', '),
		style: `margin-left:${points(x)};margin-top:${points(y)};width:${points(NOTE_WIDTH_PX)};height:${points(height)}`,
	};
}
