import { describe, expect, it } from 'vitest';
import { wordBaselinePx, wordPdfPt, wordPdfPx } from './pdf-units';

describe("Word's 1/600 inch PDF grid", () => {
	it('rounds font sizes as Word 16 writes them in Tf, ties up', () => {
		const sizes: [number, number][] = [
			[8, 8.04],
			[8.5, 8.52],
			[9, 9],
			[10, 9.96],
			[10.5, 10.56],
			[11, 11.04],
			[12, 12],
			[13, 12.96],
			[13.5, 13.56],
			[16, 15.96],
			[50, 50.04],
			[70.5, 70.56],
			[72, 72],
		];
		for (const [nominal, drawn] of sizes) expect(wordPdfPt(nominal)).toBeCloseTo(drawn, 9);
	});

	it('rounds the A4 page box to 595.32 x 841.92 pt', () => {
		expect(wordPdfPt(11906 / 20)).toBeCloseTo(595.32, 9);
		expect(wordPdfPt(16838 / 20)).toBeCloseTo(841.92, 9);
	});

	it('snaps CSS pixel lengths to 0.16 px steps', () => {
		expect(wordPdfPx((11 * 96) / 72)).toBeCloseTo((11.04 * 96) / 72, 9);
		expect(wordPdfPx(100.07)).toBeCloseTo(100.0, 9);
		expect(wordPdfPx(100.09)).toBeCloseTo(100.16, 9);
	});

	describe("Word's baselines (wordBaselinePx)", () => {
		const rows = (px: number) => Math.round((px * 600) / 96);
		// Line top, height and ascent in 1/600 inch rows; the expected rows are those of Word 16 PDFs.
		const baseline = (top: number, height: number, ascent: number) =>
			rows(wordBaselinePx((top * 96) / 600, (height * 96) / 600, (ascent * 96) / 600));

		it("hangs a line whose height has a fraction under half a row from its top's row", () => {
			// Calibri 12 pt at a 1417 twip top margin: 122.07 rows a line, 95.21 rows of ascent.
			const top = 1417 / 2.4;
			expect(
				[0, 1, 2, 3].map((line) => baseline(top + line * 122.0703, 122.0703, 95.2148)),
			).toEqual([685, 807, 930, 1052]);
		});

		it('stands a line whose height has a larger fraction on its bottom row', () => {
			// Calibri 11 pt: 111.90 rows a line, 24.62 rows of descent under the baseline.
			const top = 1417 / 2.4;
			expect(
				[0, 1, 7, 8].map((line) => baseline(top + line * 111.8978, 111.8978, 87.2803)),
			).toEqual([677, 789, 1461, 1572]);
		});
	});
});
