/**
 * The text of a morphing paragraph, run by run: which runs a stretch of it
 * spans and how far it advances, measured the way the browser lays it out.
 * Shared by the word layout (`morph-text-lines`) and its ink (`morph-text-ink`).
 *
 * @module render/morph-text-spans
 */
import type { FrameParagraph, MorphTextMeasure } from './morph-text-frame-read';

/** A rectangle in px from an element's own top-left corner. */
export interface MorphTextRect {
	left: number;
	top: number;
	right: number;
	bottom: number;
}

export function pxOf(value: string | number | undefined): number | undefined {
	if (typeof value === 'number') {
		return value;
	}
	const parsed = typeof value === 'string' ? Number.parseFloat(value) : Number.NaN;
	return Number.isFinite(parsed) ? parsed : undefined;
}

/** The runs of `paragraph` that `[start, end)` touches, with the slice of each. */
export function slices(paragraph: FrameParagraph, start: number, end: number) {
	const out: { index: number; text: string; from: number }[] = [];
	let offset = 0;
	for (const [index, run] of paragraph.runs.entries()) {
		const from = Math.max(start, offset);
		const to = Math.min(end, offset + run.text.length);
		if (to > from) {
			out.push({ index, text: run.text.slice(from - offset, to - offset), from });
		}
		offset += run.text.length;
	}
	return out;
}

/** Advance width of `paragraph.text.slice(start, end)`, measured run by run. */
export function spanWidth(
	paragraph: FrameParagraph,
	start: number,
	end: number,
	measure: MorphTextMeasure,
): number | undefined {
	let width = 0;
	for (const { index, text } of slices(paragraph, start, end)) {
		const run = paragraph.runs[index];
		const advance = measure(text, paragraph.fonts[index] ?? '', run?.style.fontKerning !== 'none');
		if (advance === undefined) {
			return undefined;
		}
		width += advance + (pxOf(run?.style.letterSpacing) ?? 0) * [...text].length;
	}
	return width;
}

/** The font family of a canvas `font` shorthand. */
export const familyOf = (font: string): string => font.replace(/^.*?[\d.]+px\s*/u, '');
