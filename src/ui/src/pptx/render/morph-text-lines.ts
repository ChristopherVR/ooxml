/**
 * Where each WORD of a morphing text box sits once its paragraphs wrap, so a
 * pair whose line breaks differ can be split into the pieces PowerPoint flies
 * on their own (see `morph-text-pieces`).
 *
 * This is a greedy word wrap over the canvas measurement `morph-text-frame`
 * already uses, which is how the browser breaks plain Latin text, followed by
 * the CSS the renderer paints with: `margin-left` and `text-indent` narrow
 * and shift a paragraph's lines, `text-align` places each line (`justify`
 * spreads every line but a paragraph's last over its spaces), a line is as
 * tall as its largest run, `margin-bottom` follows each paragraph, and the
 * flex column centres or bottom-aligns the whole block for a middle or bottom
 * anchor.
 *
 * It refuses anything it cannot reproduce with confidence: mixed font sizes in
 * different faces or under exact line spacing, distributed alignment,
 * transformed, shifted or doubled-space text, a word wider than its column,
 * any line whose break sits within a pixel of the column edge (where a
 * rounding difference could make the browser break it elsewhere), a
 * shrink-wrapped (`anchorCtr`) block, and text that overflows a box it is not
 * anchored to the top of.
 *
 * @module render/morph-text-lines
 */
import type { FrameParagraph, MorphTextMeasure, TextFrame } from './morph-text-frame-read';
import type { MorphTextInkMeasure } from './morph-text-ink';
import { glyphInk, lineBaseline, union } from './morph-text-ink';
import type { MorphTextRect } from './morph-text-spans';
import { familyOf, pxOf, slices, spanWidth } from './morph-text-spans';
import { lineHeightToPx } from './text-line-height';

export type { MorphTextRect } from './morph-text-spans';

/** One word, in px from the element's own top-left corner. */
export interface MorphTextWord {
	/** Paragraph index within the frame. */
	paragraph: number;
	/** Line index within the whole frame. */
	line: number;
	/** Left and right advance extent of the word. */
	x0: number;
	x1: number;
	/** Where its glyphs paint, when an ink measurer was given: all of them, and each. */
	ink?: MorphTextRect;
	glyphs?: MorphTextRect[];
}

/** One laid-out line. */
export interface MorphTextLine {
	top: number;
	height: number;
	/** Left edge of the line's first word and right edge of its last. */
	x0: number;
	x1: number;
	/** Whether `justify` spread the line's words apart. */
	justified: boolean;
}

/** A frame's words, in reading order, and its lines. */
export interface MorphTextLayout {
	words: MorphTextWord[];
	lines: MorphTextLine[];
	/** Bottom of the last line, plus the space after the last paragraph. */
	bottom: number;
}

/** How close (px) a break may sit to the column edge and still be trusted. */
const BREAK_MARGIN_PX = 1;
/** Upper bound on the words laid out; longer bodies keep the box interpolation. */
const MAX_WORDS = 400;
/** Sub-pixel overflow a middle or bottom anchored block may have. */
const OVERFLOW_EPSILON_PX = 0.5;

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
	availableOf: (line: number) => number,
	wraps: boolean,
	measure: MorphTextMeasure,
): number[][] | undefined {
	const lines: number[][] = [];
	let first = 0;
	while (first < spans.length) {
		const available = availableOf(lines.length);
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

/** Whether the paragraph's text paints as measured and its lines can be sized. */
function layoutable(paragraph: FrameParagraph): boolean {
	const sizes = new Set(paragraph.runs.map((run) => pxOf(run.style.fontSize) ?? paragraph.size));
	const families = new Set(paragraph.fonts.map(familyOf));
	const plain = paragraph.runs.every(
		(run) =>
			!run.style.verticalAlign &&
			(run.style.textTransform ?? 'none') === 'none' &&
			(run.style.fontVariantCaps ?? 'normal') === 'normal',
	);
	const justified = paragraph.textAlign === 'justify';
	return (
		plain &&
		(sizes.size <= 1 || (typeof paragraph.lineHeight === 'number' && families.size === 1)) &&
		(!justified || (paragraph.textAlignLast === undefined && !/ {2}| /u.test(paragraph.text)))
	);
}

/** Height and baseline (from the line's top) of the line holding `[start, end)`. */
function lineMetrics(
	paragraph: FrameParagraph,
	start: number,
	end: number,
	ink: MorphTextInkMeasure | undefined,
): { height: number; baseline: number | undefined } {
	const touched = slices(paragraph, start, end).map(({ index }) => index);
	const sizeOf = (index: number) =>
		pxOf(paragraph.runs[index]?.style.fontSize) ?? paragraph.ownStrut;
	const largest = touched.length ? Math.max(paragraph.ownStrut, ...touched.map(sizeOf)) : 0;
	const height = lineHeightToPx(largest || paragraph.strut, paragraph.lineHeight);
	const baseline = ink && lineBaseline(paragraph, touched, sizeOf, height, ink);
	return { height, baseline };
}

/**
 * Lay out `frame` word by word, or `undefined` when its layout cannot be
 * reproduced with confidence.
 *
 * @param ink - Ink measurement; without it the words carry no ink boxes.
 */
export function layoutMorphTextWords(
	frame: TextFrame,
	measure: MorphTextMeasure,
	ink?: MorphTextInkMeasure,
): MorphTextLayout | undefined {
	if (frame.anchorCenter) {
		return undefined;
	}
	const words: MorphTextWord[] = [];
	const lines: MorphTextLine[] = [];
	let top = frame.insetTop;
	for (const [index, paragraph] of frame.paragraphs.entries()) {
		const spans = wordSpans(paragraph.text);
		if (!layoutable(paragraph) || words.length + spans.length > MAX_WORDS) {
			return undefined;
		}
		const column = frame.available - paragraph.marginLeftPx;
		const indentOf = (line: number) => (line === 0 ? paragraph.textIndentPx : 0);
		const broken = spans.length
			? breakParagraph(paragraph, spans, (line) => column - indentOf(line), frame.wraps, measure)
			: [];
		if (!broken || !(column > 0)) {
			return undefined;
		}
		for (const [lineIndex, indices] of (broken.length ? broken : [[]]).entries()) {
			const lineStart = spans[indices[0] ?? 0]?.[0] ?? 0;
			const lineEnd = spans[indices.at(-1) ?? 0]?.[1] ?? 0;
			const nextStart = spans[(indices.at(-1) ?? 0) + 1]?.[0] ?? paragraph.text.length;
			const width = indices.length ? spanWidth(paragraph, lineStart, lineEnd, measure) : 0;
			const { height, baseline } = lineMetrics(paragraph, lineStart, nextStart, ink);
			if (width === undefined || !(height > 0)) {
				return undefined;
			}
			const origin = frame.insetLeft + paragraph.marginLeftPx + indentOf(lineIndex);
			const room = column - indentOf(lineIndex);
			const last = lineIndex === broken.length - 1;
			const gaps = indices.length - 1;
			const spread = paragraph.textAlign === 'justify' && frame.wraps && !last && gaps > 0;
			const extra = spread ? Math.max(0, room - width) / gaps : 0;
			const left =
				spread || frame.align === 'left'
					? origin
					: frame.align === 'right'
						? origin + room - width
						: origin + (room - width) / 2;
			for (const [k, word] of indices.entries()) {
				const [start, end] = spans[word] ?? [0, 0];
				const before = spanWidth(paragraph, lineStart, start, measure);
				const through = spanWidth(paragraph, lineStart, end, measure);
				if (before === undefined || through === undefined) {
					return undefined;
				}
				const x0 = left + before + k * extra;
				const glyphs =
					ink && baseline !== undefined
						? glyphInk(paragraph, start, end, x0, top + baseline, measure, ink)
						: undefined;
				const box = glyphs && union(glyphs);
				words.push({
					paragraph: index,
					line: lines.length,
					x0,
					x1: left + through + k * extra,
					...(glyphs && box ? { ink: box, glyphs } : {}),
				});
			}
			const x1 = left + width + gaps * extra;
			lines.push({ top, height, x0: left, x1, justified: extra > 0 });
			top += height;
		}
		top += paragraph.spaceAfterPx;
	}
	return anchored(frame, { words, lines, bottom: top });
}

/** Move a laid-out block to its vertical anchor, or `undefined` if it overflows one. */
function anchored(frame: TextFrame, layout: MorphTextLayout): MorphTextLayout | undefined {
	if (frame.vAlign === 'top') {
		return layout;
	}
	const area = frame.boxHeight - frame.insetTop - frame.insetBottom;
	const free = area - (layout.bottom - frame.insetTop);
	if (free < -OVERFLOW_EPSILON_PX) {
		return undefined;
	}
	const dy = frame.vAlign === 'bottom' ? free : free / 2;
	const shift = (rect: MorphTextRect): MorphTextRect => ({
		...rect,
		top: rect.top + dy,
		bottom: rect.bottom + dy,
	});
	return {
		words: layout.words.map((word) =>
			word.ink && word.glyphs
				? { ...word, ink: shift(word.ink), glyphs: word.glyphs.map(shift) }
				: word,
		),
		lines: layout.lines.map((line) => ({ ...line, top: line.top + dy })),
		bottom: layout.bottom + dy,
	};
}
