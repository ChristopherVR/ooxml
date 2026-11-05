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

import type { MorphFontMetrics } from './morph-text-baseline';
import {
	measureMorphFontWithCanvas,
	morphTextCapHeight,
	morphTextFirstBaseline,
} from './morph-text-baseline';
import type { MorphTextMeasure, ReadFrameOptions, TextFrame } from './morph-text-frame-read';
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
	/**
	 * The rotation (degrees) BOTH halves stand at, when it is not zero. The
	 * anchors and scales are then in the boxes' own, unrotated coordinates, and
	 * `centres` holds the two boxes' centres the rotation turns them about.
	 */
	rotation?: number;
	centres?: { from: { x: number; y: number }; to: { x: number; y: number } };
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
/** Two halves within this many degrees of each other stand at the same rotation. */
const ROTATION_EPSILON_DEG = 0.01;

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
	options: Pick<ReadFrameOptions, 'withBody'> & { metrics?: MorphFontMetrics } = {},
): MorphTextAlignment | undefined {
	const rotation = normalDegrees(from.rotation ?? 0);
	if (
		(from as { shapeType?: string }).shapeType !== (to as { shapeType?: string }).shapeType ||
		!hasTextProperties(from) ||
		!hasTextProperties(to) ||
		from.text !== to.text ||
		Math.abs(rotation - normalDegrees(to.rotation ?? 0)) > ROTATION_EPSILON_DEG
	) {
		return undefined;
	}
	const read = { rotated: true, ...(options.withBody ? { withBody: true } : {}) };
	const a = readFrame(from, read);
	const b = readFrame(to, read);
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
	const anchors = {
		from: { x: a.anchorX, y: a.anchorY },
		to: { x: b.anchorX, y: b.anchorY },
		...(rotation === 0
			? {}
			: {
					rotation,
					centres: {
						from: { x: from.x + from.width / 2, y: from.y + from.height / 2 },
						to: { x: to.x + to.width / 2, y: to.y + to.height / 2 },
					},
				}),
	};
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
	// Two faces put their baselines at different depths in a line, so the
	// halves are laid baseline on baseline (see `morph-text-baseline`).
	// One face at another size scales uniformly, which the line extents
	// already say exactly; only a change of face needs the face's metrics.
	const faceOf = (font: string) => font.replace(/[\d.]+px/gu, '');
	const sameFace = a.paragraphs.every(
		(paragraph, index) =>
			paragraph.fonts.map(faceOf).join('|') === b.paragraphs[index]?.fonts.map(faceOf).join('|'),
	);
	const metrics = options.metrics ?? measureMorphFontWithCanvas;
	const baseA = sameFace ? undefined : morphTextFirstBaseline(a, metrics);
	const baseB = sameFace ? undefined : morphTextFirstBaseline(b, metrics);
	if (baseA === undefined || baseB === undefined) {
		return { ...anchors, scaleX, scaleY };
	}
	anchors.from.y += baseA;
	anchors.to.y += baseB;
	// One line in one face each: PowerPoint stretches the glyphs onto the
	// other face's capital height (a Calibri -> Arial "L" grows 30 -> 33 -> 35
	// px through the morph, one stem, never two), about the baseline.
	const capA = morphTextCapHeight(a, metrics);
	const capB = morphTextCapHeight(b, metrics);
	return {
		...anchors,
		scaleX,
		scaleY: capA !== undefined && capB !== undefined ? capB / capA : scaleY,
	};
}

/**
 * The translation that, about a box centred on `pivot`, lays the local anchor
 * `own` (scaled by `scale`) onto the slide point `target`. `own` and `target`
 * are both read unrotated; `rotation` turns the offset the way the rendered
 * boxes are turned, which is how a rotated pair's text stays glyph over glyph.
 */
function translationOnto(
	target: { x: number; y: number },
	targetPivot: { x: number; y: number },
	own: { x: number; y: number },
	pivot: { x: number; y: number },
	sx: number,
	sy: number,
	rotation: number,
): { dx: number; dy: number } {
	const vx = target.x - targetPivot.x - sx * (own.x - pivot.x);
	const vy = target.y - targetPivot.y - sy * (own.y - pivot.y);
	const radians = (rotation * Math.PI) / 180;
	const cos = Math.cos(radians);
	const sin = Math.sin(radians);
	return {
		dx: targetPivot.x - pivot.x + vx * cos - vy * sin,
		dy: targetPivot.y - pivot.y + vx * sin + vy * cos,
	};
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
	const pivot = { x: to.x + to.width / 2, y: to.y + to.height / 2 };
	const targetPivot = alignment.rotation ? (alignment.centres?.from ?? pivot) : pivot;
	return {
		...translationOnto(
			alignment.from,
			targetPivot,
			alignment.to,
			pivot,
			sx,
			sy,
			alignment.rotation ?? 0,
		),
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
	const pivot = { x: from.x + from.width / 2, y: from.y + from.height / 2 };
	const targetPivot = alignment.rotation ? (alignment.centres?.to ?? pivot) : pivot;
	return {
		...translationOnto(
			alignment.to,
			targetPivot,
			alignment.from,
			pivot,
			alignment.scaleX,
			alignment.scaleY,
			alignment.rotation ?? 0,
		),
		sx: alignment.scaleX,
		sy: alignment.scaleY,
	};
}

/** `degrees` in [0, 360). */
function normalDegrees(degrees: number): number {
	return ((degrees % 360) + 360) % 360;
}
