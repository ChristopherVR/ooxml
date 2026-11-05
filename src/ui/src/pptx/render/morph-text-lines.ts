/**
 * Where each WORD of a morphing text box sits once its paragraphs wrap, so a
 * pair whose line breaks differ can be split into the pieces PowerPoint flies
 * on their own (see `morph-text-pieces`).
 *
 * This is a greedy word wrap over the canvas measurement `morph-text-frame`
 * already uses, which is how the browser breaks plain Latin text. It refuses
 * anything it cannot reproduce with confidence: indented or spaced paragraphs,
 * mixed font sizes on one line, a word wider than the column, and any line
 * whose break sits within a pixel of the column edge, where a rounding
 * difference could make the browser break it elsewhere.
 *
 * @module render/morph-text-lines
 */
import type { FrameParagraph, MorphTextMeasure, TextFrame } from './morph-text-frame-read';
import { lineHeightToPx } from './text-line-height';

/** One word, in px from the element's own top-left corner. */
export interface MorphTextWord {
	/** Paragraph index within the frame. */
	paragraph: number;
	/** Line index within the whole frame. */
	line: number;
	/** Left and right ink extent of the word. */
	x0: number;
	x1: number;
}

/** One laid-out line. */
export interface MorphTextLine {
	top: number;
	height: number;
	/** Left edge of the line's first word and right edge of its last. */
	x0: number;
	x1: number;
}

/** A frame's words, in reading order, and its lines. */
export interface MorphTextLayout {
	words: MorphTextWord[];
	lines: MorphTextLine[];
}

/** How close (px) a break may sit to the column edge and still be trusted. */
const BREAK_MARGIN_PX = 1;
/** Upper bound on the words laid out; longer bodies keep the box interpolation. */
const MAX_WORDS = 400;

function pxOf(value: string | number | undefined): number | undefined {
	if (typeof value === 'number') {
		return value;
	}
	const parsed = typeof value === 'string' ? Number.parseFloat(value) : Number.NaN;
	return Number.isFinite(parsed) ? parsed : undefined;
}

/** Advance width of `paragraph.text.slice(start, end)`, measured run by run. */
export function spanWidth(
	paragraph: FrameParagraph,
	start: number,
	end: number,
	measure: MorphTextMeasure,
): number | undefined {
	let offset = 0;
	let width = 0;
	for (const [index, run] of paragraph.runs.entries()) {
		const from = Math.max(start, offset);
		const to = Math.min(end, offset + run.text.length);
		offset += run.text.length;
		if (to <= from) {
			continue;
		}
		const text = run.text.slice(from - (offset - run.text.length), to - (offset - run.text.length));
		const advance = measure(text, paragraph.fonts[index] ?? '', run.style.fontKerning !== 'none');
		if (advance === undefined) {
			return undefined;
		}
		width += advance + (pxOf(run.style.letterSpacing) ?? 0) * [...text].length;
	}
	return width;
}

/** The words of one paragraph as `[start, end)` offsets; only a plain space separates them. */
function wordSpans(text: string): [number, number][] {
	const spans: [number, number][] = [];
	for (const match of text.matchAll(/[^ ]+/gu)) {
		spans.push([match.index, match.index + match[0].length]);
	}
	return spans;
}

/** Greedy breaks of one paragraph into lines of word indices, or `undefined` when unsure. */
function breakParagraph(
	paragraph: FrameParagraph,
	spans: [number, number][],
	available: number,
	wraps: boolean,
	measure: MorphTextMeasure,
): number[][] | undefined {
	const lines: number[][] = [];
	let first = 0;
	while (first < spans.length) {
		let last = first;
		const start = spans[first]?.[0] ?? 0;
		const single = spanWidth(paragraph, start, spans[first]?.[1] ?? 0, measure);
		if (single === undefined || (wraps && single > available - BREAK_MARGIN_PX)) {
			return undefined;
		}
		while (last + 1 < spans.length) {
			const width = spanWidth(paragraph, start, spans[last + 1]?.[1] ?? 0, measure);
			if (width === undefined) {
				return undefined;
			}
			if (!wraps || width <= available - BREAK_MARGIN_PX) {
				last += 1;
				continue;
			}
			if (width <= available + BREAK_MARGIN_PX) {
				return undefined;
			}
			break;
		}
		lines.push(Array.from({ length: last - first + 1 }, (_, k) => first + k));
		first = last + 1;
	}
	return lines;
}

/**
 * Lay out `frame` word by word, or `undefined` when its layout cannot be
 * reproduced with confidence. Only top-anchored text is laid out.
 */
export function layoutMorphTextWords(
	frame: TextFrame,
	measure: MorphTextMeasure,
): MorphTextLayout | undefined {
	if (frame.vAlign !== 'top') {
		return undefined;
	}
	const words: MorphTextWord[] = [];
	const lines: MorphTextLine[] = [];
	let top = frame.insetTop;
	for (const [index, paragraph] of frame.paragraphs.entries()) {
		const sizes = new Set(paragraph.runs.map((run) => pxOf(run.style.fontSize) ?? paragraph.size));
		if (paragraph.indented || sizes.size > 1) {
			return undefined;
		}
		const height = lineHeightToPx(paragraph.strut, paragraph.lineHeight);
		const spans = wordSpans(paragraph.text);
		if (words.length + spans.length > MAX_WORDS || !(height > 0)) {
			return undefined;
		}
		const broken = spans.length
			? breakParagraph(paragraph, spans, frame.available, frame.wraps, measure)
			: [];
		if (!broken) {
			return undefined;
		}
		for (const indices of broken.length ? broken : [[]]) {
			const lineStart = spans[indices[0] ?? 0]?.[0] ?? 0;
			const lineEnd = spans[indices.at(-1) ?? 0]?.[1] ?? 0;
			const width = indices.length ? spanWidth(paragraph, lineStart, lineEnd, measure) : 0;
			if (width === undefined) {
				return undefined;
			}
			const left =
				frame.align === 'left'
					? frame.insetLeft
					: frame.align === 'right'
						? frame.insetLeft + frame.available - width
						: frame.insetLeft + (frame.available - width) / 2;
			for (const word of indices) {
				const [start, end] = spans[word] ?? [0, 0];
				const before = spanWidth(paragraph, lineStart, start, measure);
				const through = spanWidth(paragraph, lineStart, end, measure);
				if (before === undefined || through === undefined) {
					return undefined;
				}
				words.push({ paragraph: index, line: lines.length, x0: left + before, x1: left + through });
			}
			lines.push({ top, height, x0: left, x1: left + width });
			top += height;
		}
		top += paragraph.spaceAfterPx;
	}
	return { words, lines };
}
