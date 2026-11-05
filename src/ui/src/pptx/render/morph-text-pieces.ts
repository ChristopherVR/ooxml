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
 * A piece is described by the same mapping `morph-text-frame` uses for a whole
 * box (`MorphTextAlignment`), plus the band of each half it shows.
 *
 * @module render/morph-text-pieces
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { hasTextProperties } from 'ooxml-core/pptx';

import type { MorphTextAlignment, MorphTextMeasure } from './morph-text-frame';
import { measureMorphTextWithCanvas, readFrame } from './morph-text-frame-read';
import type { MorphTextLayout, MorphTextWord } from './morph-text-lines';
import { layoutMorphTextWords } from './morph-text-lines';

/** A rectangle in px from an element's own top-left corner. */
export interface MorphTextBand {
	left: number;
	top: number;
	right: number;
	bottom: number;
}

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
/** How far an open band edge (a line's outer end, the first or last line) reaches. */
const OPEN_EDGE_PX = 2000;

/** The part of `layout` a run of words `[first, last]` owns, open at the outer edges. */
function bandOf(layout: MorphTextLayout, first: number, last: number): MorphTextBand {
	const words = layout.words;
	const head = words[first] as MorphTextWord;
	const tail = words[last] as MorphTextWord;
	const before = words[first - 1];
	const after = words[last + 1];
	const line = layout.lines[head.line];
	const sameLine = (word: MorphTextWord | undefined): boolean =>
		word !== undefined && word.line === head.line;
	return {
		left: sameLine(before) ? ((before as MorphTextWord).x1 + head.x0) / 2 : -OPEN_EDGE_PX,
		right: sameLine(after) ? (tail.x1 + (after as MorphTextWord).x0) / 2 : OPEN_EDGE_PX,
		top: head.line === 0 ? -OPEN_EDGE_PX : (line?.top ?? 0),
		bottom:
			head.line === layout.lines.length - 1 ? OPEN_EDGE_PX : (line?.top ?? 0) + (line?.height ?? 0),
	};
}

/**
 * Split a matched text pair into the pieces PowerPoint carries, or `undefined`
 * when it cannot be laid out with confidence (the pair then keeps its box
 * interpolation). Both halves must say exactly the same thing.
 *
 * @param from - The outgoing half.
 * @param to - The incoming half.
 * @param measure - Text measurement; defaults to a canvas.
 */
export function resolveMorphTextPieces(
	from: PptxElement,
	to: PptxElement,
	measure: MorphTextMeasure = measureMorphTextWithCanvas,
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
	const a = layoutMorphTextWords(frameA, measure);
	const b = layoutMorphTextWords(frameB, measure);
	if (!a || !b || a.words.length === 0 || a.words.length !== b.words.length) {
		return undefined;
	}
	const pieces: MorphTextPiece[] = [];
	let first = 0;
	for (let index = 0; index < a.words.length; index++) {
		const next = index + 1;
		const breaks =
			next >= a.words.length ||
			a.words[next]?.line !== a.words[index]?.line ||
			b.words[next]?.line !== b.words[index]?.line;
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
			fromBand: bandOf(a, first, index),
			toBand: bandOf(b, first, index),
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
