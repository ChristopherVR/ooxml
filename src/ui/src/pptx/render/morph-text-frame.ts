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
 *   - A box that TURNS (0 -> 40 or 60 degrees, Calibri -> Arial or 32 -> 40
 *     pt) carries one glyph-aligned copy too, turning about the middle of its
 *     text (see `morph-text-transform`).
 *   - A box flipped horizontally paints its text unmirrored, and a flip that
 *     changes between the slides leaves the text alone: it travels and
 *     restyles exactly like an unflipped one, never mirroring through zero.
 *   - A box flipped vertically paints its text upside down, and gaining or
 *     losing that flip turns the text half way round about its middle
 *     (see `morph-text-flip`).
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
import { morphTextUpright } from './morph-text-flip';
import type { MorphTextInkMeasure } from './morph-text-ink';
import { resolveMorphTextPieces } from './morph-text-pieces';
import type { MorphPoint, MorphTextAlignment } from './morph-text-transform';

export type { MorphTextMeasure } from './morph-text-frame-read';
export { measureMorphTextWithCanvas } from './morph-text-frame-read';
export type { MorphTextAlignment, MorphTextTransform } from './morph-text-transform';
export {
	morphTextGhostEnd,
	morphTextGhostStart,
	morphTextIncomingEnd,
	morphTextIncomingStart,
	morphTextTail,
} from './morph-text-transform';

/** Tolerance, in px, for any one line landing off the shared mapping. */
const LINE_DRIFT_PX = 3;
/** Sub-pixel noise allowed when comparing column widths or a line against its column. */
const FIT_EPSILON_PX = 0.5;
/** Two halves within this many degrees of each other stand at the same rotation. */
const ROTATION_EPSILON_DEG = 0.01;
/** The line pitch, in font sizes, the turning pivot estimates a text block's height with. */
const PIVOT_LINE_PITCH = 1.2;

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
 * The middle of `frame`'s text block (unrotated slide px): the point a
 * turning pair's text travels straight along. `widths` are its measured line
 * widths; without them the column's middle stands in.
 */
function textMiddle(frame: TextFrame, widths: readonly number[] | undefined): MorphPoint {
	const width = widths?.length ? Math.max(...widths) : frame.available;
	const height = frame.paragraphs.reduce(
		(sum, paragraph) => sum + paragraph.size * PIVOT_LINE_PITCH + paragraph.spaceAfterPx,
		0,
	);
	const along = { left: 1, center: 0, right: -1 }[frame.align] ?? 0;
	const down = { top: 1, middle: 0, bottom: -1 }[frame.vAlign] ?? 0;
	return { x: frame.anchorX + (along * width) / 2, y: frame.anchorY + (down * height) / 2 };
}

/** The line widths of `frame`'s filled paragraphs and their counterparts', or `undefined`. */
function measureWidths(a: TextFrame, b: TextFrame, measure: MorphTextMeasure) {
	const widthsA: number[] = [];
	const widthsB: number[] = [];
	for (const [index, paragraph] of a.paragraphs.entries()) {
		if (paragraph.text.trim() === '') {
			continue;
		}
		const other = b.paragraphs[index];
		const wa = paragraphWidth(paragraph, measure);
		const wb = other ? paragraphWidth(other, measure) : undefined;
		if (wa === undefined || wb === undefined || !(wa > 0) || !(wb > 0)) {
			return undefined;
		}
		widthsA.push(wa);
		widthsB.push(wb);
	}
	return widthsA.length === 0 ? undefined : { widthsA, widthsB };
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
	fromHalf: PptxElement,
	toHalf: PptxElement,
	measure: MorphTextMeasure = measureMorphTextWithCanvas,
	options: Pick<ReadFrameOptions, 'withBody'> & {
		metrics?: MorphFontMetrics;
		ink?: MorphTextInkMeasure;
	} = {},
): MorphTextAlignment | undefined {
	// A flipped half's text is an unflipped half's, turned 180 degrees per
	// vertical flip (see `morph-text-flip`).
	const from = morphTextUpright(fromHalf);
	const to = morphTextUpright(toHalf);
	if (!from || !to) {
		return undefined;
	}
	const fromRotation = normalDegrees(from.rotation ?? 0);
	const toRotation = normalDegrees(to.rotation ?? 0);
	if (
		(from as { shapeType?: string }).shapeType !== (to as { shapeType?: string }).shapeType ||
		!hasTextProperties(from) ||
		!hasTextProperties(to) ||
		from.text !== to.text
	) {
		return undefined;
	}
	const read = {
		rotated: true,
		flippedH: true,
		...(options.withBody ? { withBody: true } : {}),
	};
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
	const turning = Math.abs(angleBetween(fromRotation, toRotation)) > ROTATION_EPSILON_DEG;
	const centres = {
		from: { x: from.x + from.width / 2, y: from.y + from.height / 2 },
		to: { x: to.x + to.width / 2, y: to.y + to.height / 2 },
	};
	const anchors: Pick<MorphTextAlignment, 'from' | 'to' | 'rotation' | 'centres'> = {
		from: { x: a.anchorX, y: a.anchorY },
		to: { x: b.anchorX, y: b.anchorY },
		...(fromRotation === 0 || turning ? {} : { rotation: fromRotation, centres }),
	};
	// A turning pair also needs the middle of its text, which it turns about.
	const turnAbout = (alignment: Omit<MorphTextAlignment, 'turn'>, widths?: number[]) => {
		if (!turning) {
			return alignment;
		}
		const pivot = textMiddle(a, widths);
		return {
			...alignment,
			centres,
			turn: {
				from: from.rotation ?? 0,
				to: to.rotation ?? 0,
				pivots: {
					from: pivot,
					to: {
						x: alignment.to.x + alignment.scaleX * (pivot.x - alignment.from.x),
						y: alignment.to.y + alignment.scaleY * (pivot.y - alignment.from.y),
					},
				},
			},
		};
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
		const widths = turning ? measureWidths(a, b, measure)?.widthsA : undefined;
		return turnAbout({ ...anchors, scaleX: 1, scaleY: 1 }, widths);
	}
	const measured = measureWidths(a, b, measure);
	if (!measured || !singleLines(a, measured.widthsA) || !singleLines(b, measured.widthsB)) {
		return undefined;
	}
	const { widthsA, widthsB } = measured;
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
	// A change of face over several lines or in several faces stretches every
	// line and run onto its own capitals (see `morph-text-piece-caps`): no
	// one transform does that, so the pair is carried piece by piece.
	if (
		!sameFace &&
		!options.withBody &&
		!turning &&
		fromRotation === 0 &&
		(morphTextCapHeight(a, metrics) === undefined ||
			morphTextCapHeight(b, metrics) === undefined) &&
		resolveMorphTextPieces(from, to, measure, options.ink, { metrics, requireCaps: true })
	) {
		return undefined;
	}
	const baseA = sameFace ? undefined : morphTextFirstBaseline(a, metrics);
	const baseB = sameFace ? undefined : morphTextFirstBaseline(b, metrics);
	if (baseA === undefined || baseB === undefined) {
		return turnAbout({ ...anchors, scaleX, scaleY }, widthsA);
	}
	const baselines = {
		...anchors,
		from: { x: anchors.from.x, y: anchors.from.y + baseA },
		to: { x: anchors.to.x, y: anchors.to.y + baseB },
	};
	// One line in one face each: PowerPoint stretches the glyphs onto the
	// other face's capital height (a Calibri -> Arial "L" grows 30 -> 33 -> 35
	// px through the morph, one stem, never two), about the baseline. With
	// several lines or faces it stretches every run onto its own counterpart's
	// capitals and every line onto its own baseline (two lines Calibri ->
	// Arial 32 pt: each "L" grows 27 -> 31 px, the baselines stay 51-52 px
	// apart), which no one transform can: such a pair is carried piece by
	// piece above, and only where that cannot be (a turning pair, a body
	// copy, no ink) do the line extents below stand as the compromise.
	const capA = morphTextCapHeight(a, metrics);
	const capB = morphTextCapHeight(b, metrics);
	return turnAbout(
		{
			...baselines,
			scaleX,
			scaleY: capA !== undefined && capB !== undefined ? capB / capA : scaleY,
		},
		widthsA,
	);
}

/** `degrees` in [0, 360). */
function normalDegrees(degrees: number): number {
	return ((degrees % 360) + 360) % 360;
}

/** The signed difference `to - from`, in (-180, 180]. */
function angleBetween(from: number, to: number): number {
	const delta = normalDegrees(to - from);
	return delta > 180 ? delta - 360 : delta;
}
