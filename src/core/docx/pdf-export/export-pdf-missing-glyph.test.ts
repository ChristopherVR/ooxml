// A `w:sym` code with no glyph in its font is drawn as the font's missing-glyph box (issue #425).
import { inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer, type TextMeasurer } from '../layout/index';
import { syntheticTrueType } from '../../text/truetype/test-support/synthetic-truetype';
import { exportDocxPdf } from './export-pdf';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function content(
	body: string,
	fonts: { family: string; data: Uint8Array }[] = [],
	measurer: TextMeasurer = createFakeMeasurer(),
) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const { bytes } = await exportDocxPdf(model, { measurer, fonts });
	const raw = Buffer.from(bytes).toString('latin1');
	const streams: string[] = [];
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		if (/\/Filter \/FlateDecode/.test(match[1]!) && !/\/Subtype \/Image|\/Length1/.test(match[1]!))
			streams.push(inflateSync(bytes.subarray(start, start + length)).toString('latin1'));
	}
	return streams.join('\n');
}

const paragraph = (char: string, font = 'Symbol') =>
	`<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="32"/></w:rPr><w:sym w:font="${font}" w:char="${char}"/></w:r><w:r><w:t xml:space="preserve">  Option one</w:t></w:r></w:p>`;

// A measurer that gives a missing-glyph box the advance Word's layout gives it.
const boxAdvance = (em: number): TextMeasurer => {
	const fake = createFakeMeasurer();
	return {
		...fake,
		widthOf: (text, font) => (text === '' ? font.sizePx * em : fake.widthOf(text, font)),
	};
};

describe('PDF export of a symbol code without a glyph', () => {
	it('draws the glyph 0 box Word draws for Symbol, without the font file', async () => {
		// Word 16 at 16 pt: glyph 0 of Symbol, an outer rectangle 1024 x 1280 and a wall of 32 units
		// (of 2048 per em), 103 units from the origin, on an advance of 1229 units.
		const pdf = await content(paragraph('F07F'), [], boxAdvance(1229 / 2048));
		const px = 16 / 0.75;
		const [, x, y, w, h, x2, , w2, h2] =
			/([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) re\n([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) re\nf\*/
				.exec(pdf)!
				.map(Number) as number[];
		expect(w).toBeCloseTo((1024 / 2048) * px, 2);
		expect(h).toBeCloseTo((1280 / 2048) * px, 2);
		expect(x2! - x!).toBeCloseTo((32 / 2048) * px, 2);
		expect(w2).toBeCloseTo(w! - (64 / 2048) * px, 2);
		expect(h2).toBeCloseTo(h! - (64 / 2048) * px, 2);
		expect(x! - 96).toBeCloseTo((103 / 2048) * px, 2);
		expect(y).toBeGreaterThan(0);
	});

	it("draws Wingdings' own box for 0x7F", async () => {
		const pdf = await content(paragraph('F07F', 'Wingdings'), [], boxAdvance(0.5));
		const px = 16 / 0.75;
		const m = /([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) re\n[^\n]* re\nf\*/.exec(pdf)!;
		expect(Number(m[3])).toBeCloseTo((768 / 2048) * px, 2);
		expect(Number(m[4])).toBeCloseTo((1480 / 2048) * px, 2);
	});

	it('draws a box for exactly the Symbol codes Word has no glyph for', async () => {
		const gaps = [0x7f, 0x80, 0x9f, 0xa0, 0xf0, 0xff];
		const present = [0x7e, 0xa1, 0xb7, 0xef, 0xfe];
		for (const code of gaps) {
			const hex = `F0${code.toString(16).toUpperCase()}`;
			expect(await content(paragraph(hex)), hex).toContain('f*');
		}
		for (const code of present) {
			const hex = `F0${code.toString(16).toUpperCase()}`;
			expect(await content(paragraph(hex)), hex).not.toContain('f*');
		}
	});

	it('draws a glyph that exists as text, with no box (control)', async () => {
		const pdf = await content(paragraph('F0B7'));
		expect(pdf).not.toContain('f*');
		expect(pdf).toContain('\xb7');
	});

	it("draws the supplied symbol font's glyph 0", async () => {
		const font = syntheticTrueType([[0xf041, 1]], 0);
		const pdf = await content(paragraph('F07F'), [{ family: 'Symbol', data: font }]);
		expect(pdf).not.toContain('f*');
		// Glyph 0 (.notdef) is written with the font, where it used to be skipped.
		expect(pdf).toMatch(/BT \/E0 [\d.]+ Tf[^\n]*\[<0001>/);
	});
});
