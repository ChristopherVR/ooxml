/**
 * The vertical mapping of a text piece whose face changes between the halves.
 *
 * PowerPoint 16 (`CreateVideo`, 32 pt, same line breaks) stretches every run
 * onto its own counterpart's capital height, about that line's own baseline:
 * two lines Calibri -> Arial grow each "L" 27 -> 31 px while the baselines stay
 * 51-52 px apart. One transform for a box cannot do both, so a pair like that
 * is carried per line (and per face within a line), each piece scaled about its
 * own baseline by its own capital-height ratio.
 *
 * @module render/morph-text-piece-caps
 */
import type { MorphFontMetrics } from './morph-text-baseline';
import type { TextFrame } from './morph-text-frame-read';
import type { MorphTextLayout, MorphTextWord } from './morph-text-lines';
import { slices } from './morph-text-spans';

/** A font shorthand without its sizes: the face (family, weight, style). */
export const faceOf = (font: string): string => font.replace(/[\d.]+px/gu, '');

/** The canvas font the first character of `word` is painted in. */
export function wordFont(frame: TextFrame, word: MorphTextWord): string | undefined {
	const paragraph = frame.paragraphs[word.paragraph];
	if (!paragraph || !word.chars) {
		return undefined;
	}
	const first = slices(paragraph, word.chars[0], word.chars[1])[0];
	return first ? paragraph.fonts[first.index] : undefined;
}

/** One half of a piece: its frame, its word layout and the piece's first word. */
export interface PieceHalf {
	frame: TextFrame;
	layout: MorphTextLayout;
	head: MorphTextWord;
}

/**
 * Where a piece's two halves anchor vertically (their baselines, from each
 * element's top) and how tall one is against the other, or `undefined` when
 * the face does not change or nothing says (the piece then keeps its line-box
 * mapping).
 */
export function pieceCapMapping(
	a: PieceHalf,
	b: PieceHalf,
	metrics: MorphFontMetrics,
): { baselines: [number, number]; scaleY: number } | undefined {
	const fontA = wordFont(a.frame, a.head);
	const fontB = wordFont(b.frame, b.head);
	if (!fontA || !fontB || faceOf(fontA) === faceOf(fontB)) {
		return undefined;
	}
	const capA = metrics(fontA)?.capHeight;
	const capB = metrics(fontB)?.capHeight;
	const lineA = a.layout.lines[a.head.line];
	const lineB = b.layout.lines[b.head.line];
	if (!capA || !capB || lineA?.baseline === undefined || lineB?.baseline === undefined) {
		return undefined;
	}
	return {
		baselines: [lineA.top + lineA.baseline, lineB.top + lineB.baseline],
		scaleY: capB / capA,
	};
}
