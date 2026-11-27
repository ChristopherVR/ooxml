/**
 * The coordinate system the decks are drawn in: a 16:9 slide of 1280 x 720 px (96 px per inch),
 * which is what the pptx SDK takes. Type sizes are written in points and converted here.
 */

export const SLIDE_W = 1280;
export const SLIDE_H = 720;
/** Slide size in EMU (13.333 in x 7.5 in). */
export const SLIDE_CX = 12_192_000;
export const SLIDE_CY = 6_858_000;

/** Pixels to EMU. */
export function emu(px: number): number {
	return Math.round(px * 9525);
}

/** Points to the pixels the SDK's text sizes are given in. */
export function pt(points: number): number {
	return (points * 4) / 3;
}

export interface Box {
	x: number;
	y: number;
	width: number;
	height: number;
}

/** Splits a width into equal columns with a gutter; returns the left edge and width of each. */
export function columns(x: number, width: number, count: number, gutter: number): Box[] {
	const column = (width - gutter * (count - 1)) / count;
	return Array.from({ length: count }, (_, index) => ({
		x: x + index * (column + gutter),
		y: 0,
		width: column,
		height: 0,
	}));
}
