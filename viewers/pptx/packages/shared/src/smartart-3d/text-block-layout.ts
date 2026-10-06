/**
 * Pure sizing of a SmartArt 3D label canvas (no `three`, no DOM), so it can be tested directly.
 *
 * A label's canvas used to be exactly the shape's text box, so text wider than the box (a long
 * word, a font substitution that is wider than the authored one), descenders, italic overhang and
 * lines spread past the box height were cut off at the canvas edge, and swapping to a shape with a
 * smaller text box clipped more. The canvas now grows symmetrically around the block centre to
 * hold the measured text, and the plane is enlarged by the same amount, so the glyphs do not move.
 *
 * @module smartart-3d/text-block-layout
 */
import type { SmartArt3DTextBlock } from '../render/smartart-3d-types';

/** Largest canvas side requested; GPUs guarantee at least 4096 for 2D textures. */
export const MAX_TEXTURE_SIDE = 4096;
/** Supersampling bounds: at least 3x the world size, at most 6x. */
const MIN_SCALE = 3;
const MAX_SCALE = 6;
/** Ascent/descent of a line's glyphs around its centre, in font sizes. */
const LINE_HALF_EM = 0.75;
/** Padding around the measured text (italic overhang, antialiasing), in font sizes. */
const PAD_EM = 0.25;

/** Canvas pixels per world unit: dense enough for the device, never beyond the texture cap. */
export function textureScale(devicePixelRatio: number): number {
	const wanted = Math.ceil((Number.isFinite(devicePixelRatio) ? devicePixelRatio : 1) * 2);
	return Math.min(MAX_SCALE, Math.max(MIN_SCALE, wanted));
}

/** The sized label canvas: world-space plane size and the pixel size to draw it at. */
export interface TextBlockLayout {
	worldWidth: number;
	worldHeight: number;
	scale: number;
	pixelWidth: number;
	pixelHeight: number;
}

/**
 * Size the canvas for `block`. `measure(text)` is the line's width in world units at the block's
 * font. The plane is never smaller than the block's text box.
 */
export function layoutTextBlock(
	block: SmartArt3DTextBlock,
	measure: (text: string) => number,
	devicePixelRatio: number,
): TextBlockLayout {
	const fontSize = Math.max(6, block.fontSize);
	let halfWidth = Math.max(1, block.maxWidth) / 2;
	let halfHeight = Math.max(1, block.maxHeight) / 2;
	for (const line of block.lines) {
		if (!line.text) continue;
		halfWidth = Math.max(halfWidth, measure(line.text) / 2 + fontSize * PAD_EM);
		halfHeight = Math.max(halfHeight, Math.abs(line.dy) + fontSize * (LINE_HALF_EM + PAD_EM));
	}
	const worldWidth = halfWidth * 2;
	const worldHeight = halfHeight * 2;
	const scale = Math.min(
		textureScale(devicePixelRatio),
		MAX_TEXTURE_SIDE / worldWidth,
		MAX_TEXTURE_SIDE / worldHeight,
	);
	return {
		worldWidth,
		worldHeight,
		scale,
		pixelWidth: Math.max(8, Math.round(worldWidth * scale)),
		pixelHeight: Math.max(8, Math.round(worldHeight * scale)),
	};
}
