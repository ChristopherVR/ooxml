/**
 * Where a morphing text box's glyphs actually paint ink, so the band a text
 * piece is clipped to (`morph-text-pieces`) can follow its own words rather
 * than the line boxes they sit in. A descender, an italic overhang or an
 * accent reaches past its line box; a band cut at the line box slices it off,
 * one cut through the ink of the neighbouring line lets that line show twice.
 *
 * The canvas reports both the ink box of a string (`actualBoundingBox*`) and
 * the font's own ascent and descent (`fontBoundingBox*`), which is what CSS
 * places a line's baseline with.
 *
 * @module render/morph-text-ink
 */
import type { FrameParagraph, MorphTextMeasure } from './morph-text-frame-read';
import type { MorphTextRect } from './morph-text-spans';
import { familyOf, slices, spanWidth } from './morph-text-spans';

/** A string's ink and its font's ascent and descent, in px about its baseline origin. */
export interface MorphTextInkBox {
	/** Ink extent left of the origin (positive reaches left) and right of it. */
	left: number;
	right: number;
	/** Ink above and below the baseline. */
	ascent: number;
	descent: number;
	/** The font's ascent and descent, which place the baseline in its line. */
	fontAscent: number;
	fontDescent: number;
}

/** Ink of `text` painted in the canvas `font`, or `undefined` when nothing can measure. */
export type MorphTextInkMeasure = (
	text: string,
	font: string,
	kerning: boolean,
) => MorphTextInkBox | undefined;

let context: CanvasRenderingContext2D | null | undefined;

/** The default ink measurer: a 2D canvas, the one the advance widths come from. */
export const measureMorphTextInkWithCanvas: MorphTextInkMeasure = (text, font, kerning) => {
	if (context === undefined) {
		context =
			typeof document === 'undefined' ? null : document.createElement('canvas').getContext('2d');
	}
	if (!context) {
		return undefined;
	}
	context.font = font;
	context.fontKerning = kerning ? 'auto' : 'none';
	const metrics = context.measureText(text);
	const box = {
		left: metrics.actualBoundingBoxLeft,
		right: metrics.actualBoundingBoxRight,
		ascent: metrics.actualBoundingBoxAscent,
		descent: metrics.actualBoundingBoxDescent,
		fontAscent: metrics.fontBoundingBoxAscent,
		fontDescent: metrics.fontBoundingBoxDescent,
	};
	return Object.values(box).every(Number.isFinite) ? box : undefined;
};

/**
 * The baseline (px from the line's top) of a line of `paragraph` whose runs
 * are `touched`, `height` tall. CSS centres each inline box's line height on
 * its own ascent and descent; the strut is the paragraph's own font at its own
 * size, and under exact spacing (runs at `line-height: 0`) it alone counts.
 */
export function lineBaseline(
	paragraph: FrameParagraph,
	touched: number[],
	sizeOf: (index: number) => number,
	height: number,
	ink: MorphTextInkMeasure,
): number | undefined {
	const strutFont = `normal ${paragraph.ownStrut}px ${familyOf(paragraph.fonts[0] ?? '')}`;
	const boxes = [{ font: strutFont, size: paragraph.ownStrut }];
	if (typeof paragraph.lineHeight === 'number') {
		boxes.push(
			...touched.map((index) => ({ font: paragraph.fonts[index] ?? '', size: sizeOf(index) })),
		);
	}
	let baseline = 0;
	for (const box of boxes) {
		const metrics = ink('x', box.font, false);
		if (!metrics) {
			return undefined;
		}
		const own = typeof paragraph.lineHeight === 'number' ? box.size * paragraph.lineHeight : height;
		baseline = Math.max(baseline, (own + metrics.fontAscent - metrics.fontDescent) / 2);
	}
	// The paragraph's relative `top` moves its glyphs, not its line boxes.
	return baseline + (paragraph.baselineShiftPx ?? 0);
}

/** Where each glyph of `[start, end)`, advancing from `x` on `baseline`, paints ink. */
export function glyphInk(
	paragraph: FrameParagraph,
	start: number,
	end: number,
	x: number,
	baseline: number,
	measure: MorphTextMeasure,
	ink: MorphTextInkMeasure,
): MorphTextRect[] | undefined {
	const rects: MorphTextRect[] = [];
	for (const { index, text, from } of slices(paragraph, start, end)) {
		const font = paragraph.fonts[index] ?? '';
		const kerning = paragraph.runs[index]?.style.fontKerning !== 'none';
		let offset = from;
		for (const glyph of text) {
			const origin = x + (spanWidth(paragraph, start, offset, measure) ?? Number.NaN);
			const box = ink(glyph, font, kerning);
			if (!box || !Number.isFinite(origin)) {
				return undefined;
			}
			offset += glyph.length;
			if (box.right + box.left > 0 && box.ascent + box.descent > 0) {
				rects.push({
					left: origin - box.left,
					right: origin + box.right,
					top: baseline - box.ascent,
					bottom: baseline + box.descent,
				});
			}
		}
	}
	return rects;
}

/** The smallest rectangle holding every one of `rects`, or `undefined` for none. */
export function union(rects: MorphTextRect[]): MorphTextRect | undefined {
	if (rects.length === 0) {
		return undefined;
	}
	return {
		left: Math.min(...rects.map((rect) => rect.left)),
		top: Math.min(...rects.map((rect) => rect.top)),
		right: Math.max(...rects.map((rect) => rect.right)),
		bottom: Math.max(...rects.map((rect) => rect.bottom)),
	};
}
