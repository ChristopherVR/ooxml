// @vitest-environment node
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { parseTrueType } from '../../text/index';
import { pairKerning } from '../../text/truetype/kerning';
import { exportDocxPdf } from './index';
import { referenceFonts } from './__fixtures__/word-reference/reference';
import { pdfGlyphLines } from './__fixtures__/word-sizes/pdf-glyphs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const TEXT = 'AVAWAY To Ty Wa Yo LT. AVAWAY To Ty Wa Yo LT.';
const SIZE_PT = 11;

async function exportKerned(kern: boolean) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>${kern ? '<w:kern w:val="2"/>' : ''}<w:sz w:val="22"/></w:rPr><w:t>${TEXT}</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const { bytes } = await exportDocxPdf(model, {
		fonts: referenceFonts(),
		now: new Date('2026-11-28T00:00:00Z'),
	});
	return { lines: await pdfGlyphLines(bytes) };
}

/** Where the font file puts each glyph of TEXT at 11 pt, with its pair kerning (or without). */
function expectedX(kern: boolean): number[] {
	const file = referenceFonts().find((font) => font.family === 'Liberation Serif' && !font.bold)!;
	const font = parseTrueType(file.data)!;
	const pairs = kern ? pairKerning(font) : undefined;
	const x: number[] = [];
	let pen = 0;
	let previous = 0;
	for (const char of TEXT) {
		const glyph = font.glyphFor(char.codePointAt(0)!);
		if (pairs && previous) pen += (pairs(previous, glyph) / font.unitsPerEm) * SIZE_PT;
		if (char !== ' ') x.push(pen);
		pen += (font.advance(glyph) / font.unitsPerEm) * SIZE_PT;
		previous = glyph;
	}
	return x;
}

describe('exportDocxPdf with w:kern', () => {
	it('kerns each glyph pair, with the space too, at the nominal widths', async () => {
		const { lines } = await exportKerned(true);
		const line = lines[0]!;
		const want = expectedX(true);
		expect(line.x).toHaveLength(want.length);
		const origin = line.x[0]!;
		const errors = line.x.map((x, index) => x - origin - (want[index]! - want[0]!));
		// Pair kerning at the nominal widths, within the 1/1000 em rounding of the offsets.
		expect(Math.max(...errors.map(Math.abs))).toBeLessThan(0.05);
		// The space pairs (Y and space, space and T) pull the later words left of the unkerned places.
		const plain = expectedX(false);
		expect(want.at(-1)! - want[0]!).toBeLessThan(plain.at(-1)! - plain[0]! - 1);
	});

	it('matches the unkerned control when w:kern is absent', async () => {
		const { lines } = await exportKerned(false);
		const line = lines[0]!;
		const want = expectedX(false);
		const origin = line.x[0]!;
		const errors = line.x.map((x, index) => x - origin - (want[index]! - want[0]!));
		expect(Math.max(...errors.map(Math.abs))).toBeLessThan(0.05);
	});
});
