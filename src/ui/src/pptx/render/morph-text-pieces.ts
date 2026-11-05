/**
 * The pieces a rewrapped text box is carried in.
 *
 * Measured on PowerPoint 16's render (`CreateVideo`, 60 fps) of a synthetic
 * pair built through COM: one Calibri paragraph in a 500 pt box with no
 * insets, 28 pt on the first slide and 36 pt on the second, played as Morph
 * by object and as Morph by word (both carry these pieces the same way). The
 * two sizes break the paragraph differently,
 *
 *     28 pt: "Lorem ipsum dolor sit amet consectetur" / "adipiscing elit sed
 *            do eiusmod tempor" / "incididunt"
 *     36 pt: "Lorem ipsum dolor sit amet" / "consectetur adipiscing elit sed
 *            do" / "eiusmod tempor incididunt"
 *
 * and PowerPoint flies the FIVE runs of words that share a line in both
 * layouts ("Lorem ... amet", "consectetur", "adipiscing elit sed do",
 * "eiusmod tempor", "incididunt") as separate pieces. Each piece's box
 * travels in a straight line from where its words sit on the first slide to
 * where they sit on the second, and scales from one size to the other, on the
 * one morph curve: template-matched at three frames of the 960 x 540 by-word
 * render, every piece's x, y and scale read the same eased progress within
 * 0.02 (0.19, 0.49 and 0.74 at frames 85, 105 and 125 of a morph starting at
 * frame 63.5). This module's pieces, painted by the React demo, match the
 * 1280 x 720 by-object render to a mean pixel difference of 0.8-1.6/255 at
 * 10-90% (5.5/255 when the box was interpolated whole).
 *
 * The same pair rendered again with one property changed (right and centre
 * alignment, a middle or bottom anchor, a first-line or hanging indent, two
 * paragraphs with space before and after, one word at a larger size, Georgia
 * italic, and Calibri -> Cambria or 28 -> 32 pt with the box widened so the
 * breaks stay the same) flies the same runs the same way: a middle or bottom
 * anchored box flies them between its anchored lines, an indented line
 * between its indented positions, and a pair whose breaks stay the same
 * carries each line on its own scale (Calibri -> Cambria stretches its lines
 * by 1.050, 1.047 and 1.079, which no one scale of the box matches). A
 * justified paragraph is the exception: its words are spaced by however much
 * room their line has left, and mid-morph each word is sharp at its own
 * interpolated place, so every word of a justified line is a piece of its own.
 * Painted by the React demo these match PowerPoint to a mean of 1.1-1.3/255
 * at 10-90% (4.4-4.8/255 while the anchored, indented and mixed-size pairs
 * kept their box interpolation, 2.7/255 while the justified pair was laid out
 * as left-aligned).
 *
 * A piece is described by the same mapping `morph-text-frame` uses for a whole
 * box (`MorphTextAlignment`), plus the band of each half it shows
 * (`morph-text-bands`).
 *
 * @module render/morph-text-pieces
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { hasTextProperties } from 'ooxml-core/pptx';

import type { MorphTextAlignment, MorphTextMeasure } from './morph-text-frame';
import { measureMorphTextWithCanvas, readFrame } from './morph-text-frame-read';
import type { MorphTextInkMeasure } from './morph-text-ink';
import { measureMorphTextInkWithCanvas } from './morph-text-ink';
import type { MorphTextBand } from './morph-text-bands';
import { morphTextBandOf, morphTextBoundaries } from './morph-text-bands';
import type { MorphTextLayout, MorphTextWord } from './morph-text-lines';
import { layoutMorphTextWords } from './morph-text-lines';

export type { MorphTextBand } from './morph-text-bands';

/** One piece: how it maps across, and the part of each half it shows. */
export interface MorphTextPiece {
	alignment: MorphTextAlignment;
	/** What the outgoing half shows of this piece. */
	fromBand: MorphTextBand;
	/** What the incoming half shows of this piece. */
	toBand: MorphTextBand;
}

/** A pair split into pieces, and how tall each half's box must be to hold its text. */
export interface MorphTextPieceSplit {
	pieces: MorphTextPiece[];
	/**
	 * The halves' heights grown to cover their last line. A box whose text
	 * overflows it (no autofit) still paints that text on the slide, but a
	 * binding may clip a copy to its box; top-anchored text does not move when
	 * its box grows, so the copies are painted at these heights.
	 */
	fromHeight: number;
	toHeight: number;
}

/** Most pieces a pair may split into before it keeps its box interpolation. */
const MAX_PIECES = 40;
/**
 * Split a matched text pair into the pieces PowerPoint carries, or `undefined`
 * when it cannot be laid out with confidence (the pair then keeps its box
 * interpolation). Both halves must say exactly the same thing.
 *
 * @param from - The outgoing half.
 * @param to - The incoming half.
 * @param measure - Text measurement; defaults to a canvas.
 * @param ink - Ink measurement for the bands' edges; defaults to a canvas
 *   when `measure` does, and to none (line-box edges) otherwise.
 */
export function resolveMorphTextPieces(
	from: PptxElement,
	to: PptxElement,
	measure: MorphTextMeasure = measureMorphTextWithCanvas,
	ink: MorphTextInkMeasure | undefined = measure === measureMorphTextWithCanvas
		? measureMorphTextInkWithCanvas
		: undefined,
): MorphTextPieceSplit | undefined {
	if (
		(from as { shapeType?: string }).shapeType !== (to as { shapeType?: string }).shapeType ||
		!hasTextProperties(from) ||
		!hasTextProperties(to) ||
		from.text !== to.text
	) {
		return undefined;
	}
	const frameA = readFrame(from);
	const frameB = readFrame(to);
	if (
		!frameA ||
		!frameB ||
		frameA.paragraphs.length !== frameB.paragraphs.length ||
		frameA.paragraphs.some((paragraph, index) => paragraph.text !== frameB.paragraphs[index]?.text)
	) {
		return undefined;
	}
	const a = layoutMorphTextWords(frameA, measure, ink);
	const b = layoutMorphTextWords(frameB, measure, ink);
	if (!a || !b || a.words.length === 0 || a.words.length !== b.words.length) {
		return undefined;
	}
	const boundariesA = morphTextBoundaries(a);
	const boundariesB = morphTextBoundaries(b);
	const pieces: MorphTextPiece[] = [];
	let first = 0;
	for (let index = 0; index < a.words.length; index++) {
		const next = index + 1;
		// A justified line spaces its words apart by however much room it has
		// left, which the other layout does not share: each word goes alone.
		const lineA = a.words[index]?.line ?? 0;
		const lineB = b.words[index]?.line ?? 0;
		const breaks =
			next >= a.words.length ||
			a.words[next]?.line !== lineA ||
			b.words[next]?.line !== lineB ||
			a.lines[lineA]?.justified === true ||
			b.lines[lineB]?.justified === true;
		if (!breaks) {
			continue;
		}
		const headA = a.words[first] as MorphTextWord;
		const headB = b.words[first] as MorphTextWord;
		const widthA = (a.words[index] as MorphTextWord).x1 - headA.x0;
		const widthB = (b.words[index] as MorphTextWord).x1 - headB.x0;
		const heightA = a.lines[headA.line]?.height ?? 0;
		const heightB = b.lines[headB.line]?.height ?? 0;
		if (!(widthA > 0) || !(widthB > 0) || !(heightA > 0) || !(heightB > 0)) {
			return undefined;
		}
		pieces.push({
			alignment: {
				from: { x: from.x + headA.x0, y: from.y + (a.lines[headA.line]?.top ?? 0) },
				to: { x: to.x + headB.x0, y: to.y + (b.lines[headB.line]?.top ?? 0) },
				scaleX: widthB / widthA,
				scaleY: heightB / heightA,
			},
			fromBand: morphTextBandOf(a, first, index, boundariesA),
			toBand: morphTextBandOf(b, first, index, boundariesB),
		});
		if (pieces.length > MAX_PIECES) {
			return undefined;
		}
		first = next;
	}
	const bottom = (layout: MorphTextLayout): number => {
		const last = layout.lines.at(-1);
		return last ? last.top + last.height : 0;
	};
	return {
		pieces,
		fromHeight: Math.max(from.height, Math.ceil(bottom(a))),
		toHeight: Math.max(to.height, Math.ceil(bottom(b))),
	};
}
