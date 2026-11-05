/**
 * Where a carried text's first BASELINE sits and how tall its capitals are,
 * so two restyled halves can be laid baseline on baseline and cap on cap.
 *
 * PowerPoint 16 (`CreateVideo`, 720p) on a 36 pt line changing Calibri ->
 * Arial: the "L" of "Lorem" is ONE stem whose height runs 30 -> 33 -> 33 ->
 * 34 -> 35 px through the morph, its foot on one baseline; the two faces only
 * double where their advances differ, along the line. Laid line top on line
 * top at one vertical scale, ours painted two stems, 30 and 35 px tall. The
 * same holds for the change in a box standing at 20 degrees.
 *
 * Two faces can also put their baselines at different depths in a line box
 * (CSS centres each face's ascent + descent in the line height), so the
 * anchor moves to the first baseline before anything is scaled about it.
 *
 * @module render/morph-text-baseline
 */
import type { FrameParagraph, TextFrame } from './morph-text-frame-read';

/**
 * A face's ascent, descent and (when known) capital height, in px, for a canvas
 * `font` shorthand.
 */
export type MorphFontMetrics = (
	font: string,
) => { ascent: number; descent: number; capHeight?: number } | undefined;

let canvasContext: CanvasRenderingContext2D | null | undefined;

/** The size the default metrics measure at: a canvas rounds its font boxes to whole px. */
const REFERENCE_PX = 1000;

/**
 * The default metrics: the font bounding box a 2D canvas reports, which CSS
 * lays lines out with, and the height of a capital "H". Measured at
 * {@link REFERENCE_PX} and scaled down, so whole-px rounding cannot skew a
 * small size (two sizes of one face then compare at exactly their ratio).
 */
export const measureMorphFontWithCanvas: MorphFontMetrics = (font) => {
	if (canvasContext === undefined) {
		canvasContext =
			typeof document === 'undefined' ? null : document.createElement('canvas').getContext('2d');
	}
	const size = sizeOf(font);
	if (!canvasContext || size === undefined) {
		return undefined;
	}
	canvasContext.font = font.replace(/[\d.]+px/u, `${REFERENCE_PX}px`);
	const metrics = canvasContext.measureText('H');
	const scale = size / REFERENCE_PX;
	const ascent = metrics.fontBoundingBoxAscent * scale;
	const descent = metrics.fontBoundingBoxDescent * scale;
	const capHeight = metrics.actualBoundingBoxAscent * scale;
	if (!Number.isFinite(ascent) || !Number.isFinite(descent)) {
		return undefined;
	}
	return Number.isFinite(capHeight) && capHeight > 0
		? { ascent, descent, capHeight }
		: { ascent, descent };
};

function sizeOf(font: string): number | undefined {
	const size = Number.parseFloat(/([\d.]+)px/u.exec(font)?.[1] ?? '');
	return Number.isFinite(size) && size > 0 ? size : undefined;
}

/** The line height (px) of an inline box of `size` px under the paragraph's `line-height`. */
function lineHeightOf(lineHeight: number | string, size: number): number | undefined {
	if (typeof lineHeight === 'number') {
		return lineHeight * size;
	}
	const value = Number.parseFloat(lineHeight);
	if (!Number.isFinite(value)) {
		return undefined;
	}
	if (lineHeight.trim().endsWith('px')) {
		return value;
	}
	return /^[\d.]+$/u.test(lineHeight.trim()) ? value * size : undefined;
}

/** A paragraph's one line box: how far it reaches above and below its baseline. */
function lineBox(
	paragraph: FrameParagraph,
	metrics: MorphFontMetrics,
): { above: number; below: number } | undefined {
	const fonts = [paragraph.strutFont, ...paragraph.fonts.filter((_, i) => paragraph.runs[i]?.text)];
	let above = 0;
	let below = 0;
	for (const font of fonts) {
		const size = sizeOf(font);
		const face = metrics(font);
		const height = size === undefined ? undefined : lineHeightOf(paragraph.lineHeight, size);
		if (!face || height === undefined) {
			return undefined;
		}
		const top = (height - face.ascent - face.descent) / 2 + face.ascent;
		above = Math.max(above, top);
		below = Math.max(below, height - top);
	}
	return { above, below };
}

/**
 * How far below the frame's anchor its first baseline sits, in px, or
 * `undefined` when it cannot be worked out (a `line-height` of `normal`, a
 * space before a paragraph of a block not anchored at its top, no metrics).
 * Every paragraph is taken to be one line, as the carried-text mapping needs.
 */
export function morphTextFirstBaseline(
	frame: TextFrame,
	metrics: MorphFontMetrics,
): number | undefined {
	const boxes = frame.paragraphs.map((paragraph) => lineBox(paragraph, metrics));
	const first = boxes[0];
	if (!first || boxes.some((box) => !box)) {
		return undefined;
	}
	if (frame.vAlign === 'top') {
		return frame.paragraphs[0]?.indented ? undefined : first.above;
	}
	if (frame.paragraphs.some((paragraph) => paragraph.indented)) {
		return undefined;
	}
	const height = frame.paragraphs.reduce(
		(sum, paragraph, index) =>
			sum + (boxes[index]!.above + boxes[index]!.below) + paragraph.spaceAfterPx,
		0,
	);
	return first.above - (frame.vAlign === 'bottom' ? height : height / 2);
}

/**
 * The capital height (px) a one-line, one-face frame paints, or `undefined`
 * when the frame has several lines or faces, or no metrics say.
 */
export function morphTextCapHeight(
	frame: TextFrame,
	metrics: MorphFontMetrics,
): number | undefined {
	const filled = frame.paragraphs.filter((paragraph) => paragraph.text.trim() !== '');
	const fonts = new Set(
		filled.flatMap((paragraph) =>
			paragraph.fonts.filter((_, i) => paragraph.runs[i]?.text?.trim()),
		),
	);
	if (filled.length !== 1 || frame.paragraphs.length !== 1 || fonts.size !== 1) {
		return undefined;
	}
	const capHeight = metrics([...fonts][0]!)?.capHeight;
	return capHeight !== undefined && capHeight > 0 ? capHeight : undefined;
}
