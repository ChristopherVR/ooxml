/**
 * Where each glyph of a one-line text box sits, so a pair that turns while its
 * type changes can be carried glyph by glyph (see `morph-text-glyph-plan`).
 *
 * @module render/morph-text-glyphs
 */
import type { PptxElement } from 'ooxml-core/pptx';

import type { MorphTextMeasure } from './morph-text-frame-read';
import { readFrame } from './morph-text-frame-read';
import type { MorphTextInkMeasure } from './morph-text-ink';
import { layoutMorphTextWords } from './morph-text-lines';
import { spanWidth } from './morph-text-spans';

/** One glyph's advance cell, in px from the element's own left edge. */
export interface MorphGlyphCell {
	/** Index of the glyph in the paragraph's text. */
	index: number;
	x0: number;
	x1: number;
	/** The middle of the glyph's ink (its advance cell and line box where nothing measures ink). */
	cx: number;
	cy: number;
}

/** Most glyphs a pair may be split into before it keeps its whole-box turn. */
export const MAX_MORPH_GLYPHS = 48;

/** Whether `text` holds a code unit whose glyph cannot be told apart by its offset. */
function complex(text: string): boolean {
	for (let index = 0; index < text.length; index++) {
		const code = text.charCodeAt(index);
		if (
			(code >= 0xd800 && code <= 0xdfff) ||
			(code >= 0x0300 && code <= 0x036f) ||
			(code >= 0x200b && code <= 0x200f) ||
			code === 0x2028 ||
			code === 0x2029 ||
			code < 0x20
		) {
			return true;
		}
	}
	return false;
}

/**
 * The advance cells of the visible glyphs of `element`, or `undefined` when it
 * is not one paragraph on one line of plain text.
 */
export function morphTextGlyphCells(
	element: PptxElement,
	measure: MorphTextMeasure,
	ink?: MorphTextInkMeasure,
): MorphGlyphCell[] | undefined {
	const frame = readFrame(element, { rotated: true });
	const paragraph = frame?.paragraphs[0];
	if (!frame || !paragraph || frame.paragraphs.length !== 1 || complex(paragraph.text)) {
		return undefined;
	}
	const layout = layoutMorphTextWords(frame, measure, ink);
	if (!layout || layout.lines.length !== 1) {
		return undefined;
	}
	const spans = [...paragraph.text.matchAll(/[^ ]+/gu)].map(
		(match) => [match.index, match.index + match[0].length] as const,
	);
	if (spans.length !== layout.words.length) {
		return undefined;
	}
	const cells: MorphGlyphCell[] = [];
	for (const [word, [start, end]] of spans.entries()) {
		const x0 = layout.words[word]?.x0;
		const line = layout.lines[0];
		const rects = layout.words[word]?.glyphs;
		if (x0 === undefined || !line) {
			return undefined;
		}
		const inked = rects?.length === end - start ? rects : undefined;
		for (let index = start; index < end; index++) {
			const before = spanWidth(paragraph, start, index, measure);
			const through = spanWidth(paragraph, start, index + 1, measure);
			if (before === undefined || through === undefined) {
				return undefined;
			}
			const rect = inked?.[index - start];
			cells.push({
				index,
				x0: x0 + before,
				x1: x0 + through,
				cx: rect ? (rect.left + rect.right) / 2 : x0 + (before + through) / 2,
				cy: rect ? (rect.top + rect.bottom) / 2 : line.top + line.height / 2,
			});
		}
	}
	return cells.length > 0 && cells.length <= MAX_MORPH_GLYPHS ? cells : undefined;
}
