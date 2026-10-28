/**
 * Unwarped layout of a WordArt envelope text block, the input to the
 * envelope mapping (`text-warp-envelope-map.ts`). Measured against
 * PowerPoint COM renders (2026-09-24, see `docs/guide/visual-effects.md`):
 *
 *  - Paragraphs stack as ordinary lines: baseline to baseline is
 *    `1.2 * ((1 - k) * S_prev + k * S_next)` for the lines' largest font
 *    sizes `S` ({@link LINE_PITCH_EM}, {@link BASELINE_FRACTION}); the 1.2em
 *    pitch held for all seven fonts measured (Arial, Times New Roman,
 *    Calibri, Georgia, Verdana, Segoe UI, Cambria, each within 0.3%).
 *  - With a wrap width (`a:bodyPr/@wrap` other than `none`, see
 *    `text-warp-envelope-wrap.ts`), each paragraph is first broken into rows
 *    at word boundaries; rows stack exactly like paragraphs do.
 *  - Lines are aligned against the widest line, and are NOT each stretched
 *    to the box: a short line keeps its natural width inside the block.
 *  - The block box is the union of every glyph's INK box vertically, and
 *    horizontally the union of ink boxes plus the full advance box of
 *    whitespace (a leading or trailing space widens the block).
 */
import type { EnvelopeBlockBox } from './text-warp-envelope-map';
import {
	envelopeFontSizePx,
	measureGlyphAdvances,
	measureGlyphInk,
} from './text-warp-envelope-measure';
import type {
	EnvelopeAlign,
	EnvelopeFontSpec,
	EnvelopeSegmentInput,
	GlyphOutlineLookup,
} from './text-warp-envelope-types';
import type { GlyphOutlineCommand } from './text-warp-glyph-outline';
import { outlineBounds } from './text-warp-glyph-outline';
import { wrapEnvelopeParagraph } from './text-warp-envelope-wrap';
import type { EnvelopeWrapItem } from './text-warp-envelope-wrap';

/** PowerPoint's single-spaced line pitch, in ems of the line's largest font. */
export const LINE_PITCH_EM = 1.2;

/** Where the baseline sits inside a line's pitch (fraction from its top). */
const BASELINE_FRACTION = 0.805;

/** Ink height estimate (ems above the baseline) with no outline or DOM metrics. */
const FALLBACK_CAP_HEIGHT_EM = 0.72;

const WHITESPACE_RE = /^\s$/u;

/** One glyph placed in the unwarped block. */
export interface EnvelopeBlockGlyph {
	char: string;
	segmentIndex: number;
	font: EnvelopeFontSpec;
	/** Left advance edge (layout units). */
	x: number;
	/** Baseline (layout units). */
	baseline: number;
	advance: number;
	/** The glyph's outline at `(x, baseline)`, when obtainable. */
	outline?: GlyphOutlineCommand[];
	/** Ink box; `undefined` for whitespace. */
	ink?: EnvelopeBlockBox;
}

/** The unwarped block: glyphs per paragraph plus the box the warp normalises by. */
export interface EnvelopeBlockLayout {
	/** Every paragraph's glyphs, in order (all of its wrapped rows together). */
	lines: EnvelopeBlockGlyph[][];
	/** How many rows each paragraph wrapped into (1 when it did not wrap). */
	rowCounts: number[];
	box: EnvelopeBlockBox | undefined;
}

function alignOffset(align: EnvelopeAlign, blockWidth: number, lineWidth: number): number {
	if (align === 'right') {
		return blockWidth - lineWidth;
	}
	if (align === 'center' || align === undefined) {
		return (blockWidth - lineWidth) / 2;
	}
	return 0;
}

function glyphInk(
	char: string,
	font: EnvelopeFontSpec,
	x: number,
	baseline: number,
	advance: number,
	outline: GlyphOutlineCommand[] | undefined,
): EnvelopeBlockBox | undefined {
	if (WHITESPACE_RE.test(char)) {
		return undefined;
	}
	if (outline) {
		return outlineBounds(outline);
	}
	const ink = measureGlyphInk(char, font);
	if (ink) {
		return ink.right > ink.left || ink.ascent + ink.descent > 0
			? {
					left: x + ink.left,
					right: x + ink.right,
					top: baseline - ink.ascent,
					bottom: baseline + ink.descent,
				}
			: undefined;
	}
	return {
		left: x,
		right: x + advance,
		top: baseline - envelopeFontSizePx(font) * FALLBACK_CAP_HEIGHT_EM,
		bottom: baseline,
	};
}

function lineFontSize(segments: EnvelopeSegmentInput[], fallback: number): number {
	const sizes = segments.map((s) => envelopeFontSizePx(s.font));
	return sizes.length > 0 ? Math.max(...sizes) : fallback;
}

/** One row of the unwarped block: a whole paragraph, or one wrapped row of it. */
interface EnvelopeBlockRow {
	paragraphIndex: number;
	items: EnvelopeWrapItem[];
	width: number;
	/** The row's largest font size (px), which sets its line pitch. */
	fontSize: number;
}

function measureParagraph(segments: EnvelopeSegmentInput[]): EnvelopeWrapItem[] {
	return segments.flatMap((seg) => {
		const advances = measureGlyphAdvances(seg.text, seg.font);
		return [...seg.text].map((char, i) => ({ seg, char, advance: advances[i] ?? 0 }));
	});
}

function blockRows(
	paragraphs: EnvelopeSegmentInput[][],
	wrapWidth: number | undefined,
): EnvelopeBlockRow[] {
	const rows: EnvelopeBlockRow[] = [];
	let previousSize = 0;
	paragraphs.forEach((segments, paragraphIndex) => {
		const wrapped = wrapEnvelopeParagraph(measureParagraph(segments), wrapWidth);
		// An unwrapped paragraph keeps sizing its line by every run it holds.
		const paragraphSize = lineFontSize(segments, previousSize);
		for (const items of wrapped) {
			const fontSize =
				wrapped.length === 1 || items.length === 0
					? paragraphSize
					: Math.max(...items.map((item) => envelopeFontSizePx(item.seg.font)));
			const width = items.reduce((sum, item) => sum + item.advance, 0);
			rows.push({ paragraphIndex, items, width, fontSize });
			previousSize = fontSize;
		}
	});
	return rows;
}

/**
 * Lay out `paragraphs` (each a list of styled runs) unwarped, and measure
 * the block box. The box is `undefined` when nothing on any line has ink.
 * `wrapWidth` (see `envelopeWrapWidth`) breaks each paragraph into rows no
 * wider than it at word boundaries; `undefined` keeps every paragraph on one
 * row.
 */
export function layoutEnvelopeBlock(
	paragraphs: EnvelopeSegmentInput[][],
	align: EnvelopeAlign,
	getGlyphOutline?: GlyphOutlineLookup,
	wrapWidth?: number,
): EnvelopeBlockLayout {
	const rows = blockRows(paragraphs, wrapWidth);
	const blockWidth = rows.length > 0 ? Math.max(...rows.map((row) => row.width)) : 0;

	let baseline = 0;
	let previousSize: number | undefined;
	const box = { left: Infinity, top: Infinity, right: -Infinity, bottom: -Infinity };
	let hasInk = false;
	const lines = paragraphs.map((): EnvelopeBlockGlyph[] => []);
	const rowCounts = paragraphs.map(() => 0);
	for (const row of rows) {
		const size = row.fontSize;
		if (previousSize !== undefined) {
			baseline +=
				LINE_PITCH_EM * ((1 - BASELINE_FRACTION) * previousSize + BASELINE_FRACTION * size);
		}
		previousSize = size;
		let x = alignOffset(effectiveAlign(align), blockWidth, row.width);
		const glyphs = lines[row.paragraphIndex];
		rowCounts[row.paragraphIndex]++;
		for (const { seg, char, advance } of row.items) {
			const outline = WHITESPACE_RE.test(char)
				? undefined
				: getGlyphOutline?.(char, seg.font, x, baseline);
			const usable = outline && outline.length > 0 ? outline : undefined;
			const ink = glyphInk(char, seg.font, x, baseline, advance, usable);
			if (ink) {
				hasInk = true;
				box.left = Math.min(box.left, ink.left);
				box.right = Math.max(box.right, ink.right);
				box.top = Math.min(box.top, ink.top);
				box.bottom = Math.max(box.bottom, ink.bottom);
			} else if (WHITESPACE_RE.test(char)) {
				box.left = Math.min(box.left, x);
				box.right = Math.max(box.right, x + advance);
			}
			glyphs.push({
				char,
				segmentIndex: seg.segmentIndex,
				font: seg.font,
				x,
				baseline,
				advance,
				outline: usable,
				ink,
			});
			x += advance;
		}
	}
	return { lines, rowCounts, box: hasInk ? box : undefined };
}

/** Distributed/justified alignments lay each row out left-aligned. */
function effectiveAlign(align: EnvelopeAlign): EnvelopeAlign {
	return align === 'justify' || align === 'justLow' || align === 'dist' || align === 'thaiDist'
		? 'left'
		: align;
}
