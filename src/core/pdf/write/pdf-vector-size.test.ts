import { describe, expect, it } from 'vitest';
import type { PdfCidFont } from './pdf-cid-font';
import { PdfVectorPage, sizeCorrections } from './pdf-vector';

/** A CID font whose glyphs are 500/1000 em wide, except a 278 space. */
const font: PdfCidFont = {
	family: 'Fixture',
	bold: false,
	italic: false,
	width: (text) => [...text].reduce((sum, char) => sum + (char === ' ' ? 0.278 : 0.5), 0),
	encode: (text) =>
		[...text].map((char) => char.charCodeAt(0).toString(16).padStart(4, '0')).join(''),
	build: () => {
		throw new Error('unused');
	},
};

describe('drawing embedded text at another size than its advances', () => {
	it('keeps every glyph within half a thousandth of an em of its nominal position', () => {
		const widths = [500, 278, 612, 444, 500, 229, 333, 500, 500, 278];
		const size = 14.666667;
		const drawSize = 14.72;
		const offsets = sizeCorrections(widths, size, drawSize);
		expect(offsets.every(Number.isInteger)).toBe(true);
		let drawn = 0;
		let nominal = 0;
		widths.forEach((width, index) => {
			drawn += ((width - offsets[index]!) * drawSize) / 1000;
			nominal += (width * size) / 1000;
			expect(Math.abs(drawn - nominal)).toBeLessThanOrEqual((0.5 * drawSize) / 1000 + 1e-9);
		});
	});

	it('writes the drawn size in Tf and corrects the glyphs with TJ offsets', () => {
		const page = new PdfVectorPage(100, 100);
		page.embeddedText(font, 10, 0, 0, 'aaaa aaaa', { drawSize: 9.96 });
		const op = page.content;
		expect(op).toContain('/E0 9.96 Tf');
		const tj = /\[(.*)\] TJ/.exec(op)![1]!;
		const offsets = [...tj.matchAll(/> (-?\d+)/g)].map((match) => Number(match[1]));
		// 9 glyphs, 4.278 em at 10 against 9.96: 0.17 em = 171 thousandths of the drawn size.
		expect(offsets.reduce((sum, value) => sum + value, 0)).toBe(-Math.round((4278 * 0.04) / 9.96));
	});

	it('adds only the word spacing without a drawn size', () => {
		const page = new PdfVectorPage(100, 100);
		page.embeddedText(font, 10, 0, 0, 'aa aa', { word: 2 });
		expect(page.content).toContain('/E0 10 Tf');
		expect(page.content).toContain('[<006100610020> -200 <00610061>] TJ');
	});
});
