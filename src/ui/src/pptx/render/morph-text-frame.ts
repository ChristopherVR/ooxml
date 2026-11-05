/**
 * Where a morphing text box's TEXT sits, so a matched pair can carry its
 * glyphs rather than its box.
 *
 * A matched pair normally interpolates its whole box: the incoming half starts
 * scaled onto the outgoing box and the ghost ends scaled onto the incoming one.
 * For a shape that is the morph PowerPoint plays. For a text box it is not,
 * because PowerPoint morphs the text's own line fragments, not the container.
 * Measured on PowerPoint 16's render (`CreateVideo`, 60 fps, 960 x 540):
 *
 *   - Calibri -> Courier New, same box, 32 pt: the line's right edge travels
 *     565 -> 776 px on the morph curve while its left edge stays at 49-50 px,
 *     and BOTH renders are stretched onto that interpolated line: mid-morph the
 *     Calibri word "Lorem" is visibly widened and the Courier one narrowed, and
 *     where the two faces place a letter differently (monospace against
 *     proportional) the two copies show through each other.
 *   - Calibri 24 -> 48 pt, same box: one sharp copy that scales uniformly about
 *     the line's start, 209 -> 419 px wide and 19 -> 40 px tall.
 *   - Calibri Light -> Calibri (issue #141, synthetic and the corporate deck's
 *     slides 6 -> 7): the two faces have nearly proportional advances, so the
 *     stretched renders coincide and read as one copy whose weight thickens.
 *   - A wrapped paragraph whose line breaks differ (28 -> 36 pt in a 500 pt
 *     box) splits into the pieces that share a line in BOTH layouts
 *     ("consectetur" leaves the end of line 1 for the start of line 2), and
 *     each piece flies on its own.
 *
 * The overlay can only move a whole element, so this resolves the case where
 * every paragraph is ONE line in both layouts (or the two layouts are the same
 * text in the same fonts in an equally wide column, which breaks identically):
 * then each half maps onto the other by one scale about the text's alignment
 * anchor, and the result is exact for one paragraph and close for several.
 * Wrapped text whose breaks may differ is split into its pieces instead (see
 * `morph-text-pieces`), or left to the box interpolation when it cannot be.
 *
 * @module render/morph-text-frame
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { hasTextProperties } from 'ooxml-core/pptx';

import type { MorphTextMeasure, TextFrame } from './morph-text-frame-read';
import { measureMorphTextWithCanvas, paragraphWidth, readFrame } from './morph-text-frame-read';

export type { MorphTextMeasure } from './morph-text-frame-read';
export { measureMorphTextWithCanvas } from './morph-text-frame-read';

/**
 * How the outgoing half's text maps onto the incoming half's: a point at
 * `from + v` (slide px) corresponds to `to + scale * v`.
 */
export interface MorphTextAlignment {
	from: { x: number; y: number };
	to: { x: number; y: number };
	scaleX: number;
	scaleY: number;
}

/** A `translate(dx, dy) scale(sx, sy)` about the element's own centre. */
export interface MorphTextTransform {
	dx: number;
	dy: number;
	sx: number;
	sy: number;
}

/** Tolerance, in px, for any one line landing off the shared mapping. */
const LINE_DRIFT_PX = 3;
/** Sub-pixel noise allowed when comparing column widths or a line against its column. */
const FIT_EPSILON_PX = 0.5;

/** Whether every paragraph of `frame` is laid out on one line at `widths`. */
function singleLines(frame: TextFrame, widths: number[]): boolean {
	return !frame.wraps || widths.every((width) => width <= frame.available + FIT_EPSILON_PX);
}

/** Whether a uniform `scale` maps every `from` extent onto its `to` within tolerance. */
function consistent(from: number[], to: number[], scale: number): boolean {
	return from.every((value, index) => Math.abs((to[index] ?? 0) - scale * value) <= LINE_DRIFT_PX);
}

function cumulative(values: number[]): number[] {
	let sum = 0;
	return values.map((value) => (sum += value));
}

/**
 * Resolve how a matched text pair's outgoing text maps onto its incoming text,
 * or `undefined` when the pair has to keep interpolating its box.
 *
 * @param from - The outgoing half.
 * @param to - The incoming half.
 * @param measure - Text measurement; defaults to a canvas. Without one only a
 *   pair in identical fonts and an equally wide column can be aligned.
 */
export function resolveMorphTextAlignment(
	from: PptxElement,
	to: PptxElement,
	measure: MorphTextMeasure = measureMorphTextWithCanvas,
): MorphTextAlignment | undefined {
	if (
		(from as { shapeType?: string }).shapeType !== (to as { shapeType?: string }).shapeType ||
		!hasTextProperties(from) ||
		!hasTextProperties(to) ||
		from.text !== to.text
	) {
		return undefined;
	}
	const a = readFrame(from);
	const b = readFrame(to);
	if (
		!a ||
		!b ||
		a.align !== b.align ||
		a.vAlign !== b.vAlign ||
		a.paragraphs.length !== b.paragraphs.length ||
		a.paragraphs.some((paragraph, index) => paragraph.text !== b.paragraphs[index]?.text)
	) {
		return undefined;
	}
	const anchors = { from: { x: a.anchorX, y: a.anchorY }, to: { x: b.anchorX, y: b.anchorY } };
	const sameFonts = a.paragraphs.every(
		(paragraph, index) =>
			paragraph.fonts.join('|') === b.paragraphs[index]?.fonts.join('|') &&
			paragraph.runs.map((run) => run.style.letterSpacing ?? '').join('|') ===
				b.paragraphs[index]?.runs.map((run) => run.style.letterSpacing ?? '').join('|'),
	);
	// The same text in the same fonts in columns of the same width breaks
	// identically, so it moves rigidly with its anchor and nothing need measure.
	if (sameFonts && a.wraps === b.wraps && Math.abs(a.available - b.available) <= FIT_EPSILON_PX) {
		return { ...anchors, scaleX: 1, scaleY: 1 };
	}
	const filled = a.paragraphs
		.map((paragraph, index) => ({ paragraph, other: b.paragraphs[index] }))
		.filter(({ paragraph }) => paragraph.text.trim() !== '');
	const widthsA: number[] = [];
	const widthsB: number[] = [];
	for (const { paragraph, other } of filled) {
		const wa = paragraphWidth(paragraph, measure);
		const wb = other ? paragraphWidth(other, measure) : undefined;
		if (wa === undefined || wb === undefined || !(wa > 0) || !(wb > 0)) {
			return undefined;
		}
		widthsA.push(wa);
		widthsB.push(wb);
	}
	if (widthsA.length === 0 || !singleLines(a, widthsA) || !singleLines(b, widthsB)) {
		return undefined;
	}
	const scaleX = widthsB.reduce((sum, w) => sum + w, 0) / widthsA.reduce((sum, w) => sum + w, 0);
	const linesA = a.paragraphs.map((paragraph) => paragraph.size);
	const linesB = b.paragraphs.map((paragraph) => paragraph.size);
	const scaleY = linesB.reduce((sum, h) => sum + h, 0) / linesA.reduce((sum, h) => sum + h, 0);
	if (
		!consistent(widthsA, widthsB, scaleX) ||
		!consistent(cumulative(linesA), cumulative(linesB), scaleY)
	) {
		return undefined;
	}
	return { ...anchors, scaleX, scaleY };
}

/**
 * The incoming half's STARTING transform: it is painted at its own geometry and
 * has to begin with its text laid over the outgoing text.
 */
export function morphTextIncomingStart(
	alignment: MorphTextAlignment,
	to: PptxElement,
): MorphTextTransform {
	const sx = 1 / alignment.scaleX;
	const sy = 1 / alignment.scaleY;
	const cx = to.x + to.width / 2;
	const cy = to.y + to.height / 2;
	return {
		dx: alignment.from.x - cx - sx * (alignment.to.x - cx),
		dy: alignment.from.y - cy - sy * (alignment.to.y - cy),
		sx,
		sy,
	};
}

/**
 * The ghost's ENDING transform: it is painted at the outgoing geometry and has
 * to end with its text laid over the incoming text.
 */
export function morphTextGhostEnd(
	alignment: MorphTextAlignment,
	from: PptxElement,
): MorphTextTransform {
	const cx = from.x + from.width / 2;
	const cy = from.y + from.height / 2;
	return {
		dx: alignment.to.x - cx - alignment.scaleX * (alignment.from.x - cx),
		dy: alignment.to.y - cy - alignment.scaleY * (alignment.from.y - cy),
		sx: alignment.scaleX,
		sy: alignment.scaleY,
	};
}
