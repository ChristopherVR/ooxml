// A `w:sym` code with no glyph in its font is drawn as the font's missing-glyph box (issue #425).
import { inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer } from '../layout/index';
import { syntheticTrueType } from '../../text/truetype/test-support/synthetic-truetype';
import { exportDocxPdf } from './export-pdf';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function content(body: string, fonts: { family: string; data: Uint8Array }[] = []) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const { bytes } = await exportDocxPdf(model, { measurer: createFakeMeasurer(), fonts });
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

const paragraph = (char: string) =>
	`<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="32"/></w:rPr><w:sym w:font="Symbol" w:char="${char}"/></w:r><w:r><w:t xml:space="preserve">  Option one</w:t></w:r></w:p>`;

describe('PDF export of a symbol code without a glyph', () => {
	it('draws the missing-glyph box without the font file', async () => {
		const pdf = await content(paragraph('F07F'));
		// An outlined box: two rectangles filled even-odd.
		expect(pdf).toMatch(/re\n[^\n]* re\nf\*/);
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
