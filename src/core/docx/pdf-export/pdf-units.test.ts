import { describe, expect, it } from 'vitest';
import { wordPdfPt, wordPdfPx } from './pdf-units';

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
});
