/**
 * Word's PDF grid: Word 16's "Save as PDF" writes font sizes, the page box and baselines in whole
 * 1/600 inch (0.12 pt) units, measured from its own PDFs (issue #382). DOM-free.
 */

/** `pt` rounded to the nearest 1/600 inch, ties up: 11 pt is drawn at 11.04, 10.5 at 10.56. */
export function wordPdfPt(pt: number): number {
	// The epsilon keeps exact ties (10.5 pt is 87.5 units) from rounding down after a product.
	return (Math.floor((pt * 600) / 72 + 0.5 + 1e-9) * 72) / 600;
}

/** A CSS pixel length (96 per inch) rounded to the nearest 1/600 inch, ties up. */
export function wordPdfPx(px: number): number {
	return (wordPdfPt((px * 72) / 96) * 96) / 72;
}

/** 1/600 inch device rows per CSS pixel. */
const ROWS_PER_PX = 600 / 96;
const row = (value: number) => Math.floor(value + 0.5 + 1e-9);

/**
 * Word's baseline, in CSS pixels, for a single-spaced line whose top is `topPx` from the page's
 * top, `heightPx` tall, with its baseline `ascentPx` below its top. Measured from Word 16's Save
 * as PDF (issue #382): lines keep the exact layout pitch, and the baseline is a whole number of
 * 1/600 inch rows below the line's top rounded to a row when the line's height in rows has a
 * fraction under one half, else a whole number of rows above its rounded bottom. The ascent rows
 * (the larger of `floor(ascent + 0.29)` and the height's whole rows less `floor(descent + 0.45)`)
 * and the descent rows (`descent` rounded) are fitted to Word's PDFs at 6.5 to 30 pt; checked on
 * single-spaced text, other spacing uses the same rule.
 */
export function wordBaselinePx(topPx: number, heightPx: number, ascentPx: number): number {
	const top = topPx * ROWS_PER_PX;
	const height = heightPx * ROWS_PER_PX;
	const ascent = ascentPx * ROWS_PER_PX;
	const descent = height - ascent;
	const rows =
		height - Math.floor(height) < 0.5
			? row(top) +
				Math.max(Math.floor(ascent + 0.29), Math.floor(height) - Math.floor(descent + 0.45))
			: row(top + height) - row(descent);
	return rows / ROWS_PER_PX;
}
