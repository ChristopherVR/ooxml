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
