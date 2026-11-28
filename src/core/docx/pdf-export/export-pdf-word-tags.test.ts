// @vitest-environment node
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { exportDocxPdf } from './index';
import { referenceFonts } from './__fixtures__/word-reference/reference';
import { pdfGlyphLines, type PdfGlyphLine } from './__fixtures__/word-sizes/pdf-glyphs';
import reference from './__fixtures__/word-tags/word-tags.json';

/** Word 16.0.20625's Save as PDF of each document, as `extract-word.ts` read it. */
const WORD: Record<keyof typeof reference, PdfGlyphLine[]> = reference;

async function exported(name: string): Promise<PdfGlyphLine[]> {
	const docx = readFileSync(new URL(`./__fixtures__/word-tags/${name}.docx`, import.meta.url));
	const { model } = await loadDocx(new Uint8Array(docx));
	const { bytes } = await exportDocxPdf(model, {
		fonts: referenceFonts(),
		now: new Date('2026-11-28T00:00:00Z'),
	});
	return pdfGlyphLines(bytes);
}

describe('tags after an untagged content control (#415)', async () => {
	const ours = await exported('issue-415-tags-after-untagged-control');
	const word = WORD['issue-415'];

	it("breaks all 104 paragraphs' lines where Word does", () => {
		expect(ours.map(({ page, text }) => ({ page, text }))).toEqual(
			word.map(({ page, text }) => ({ page, text })),
		);
	});

	it("puts each text glyph within 3 pt of Word's and ends the line like Word", () => {
		const errors = word.flatMap((line, index) =>
			line.x.map((x, glyph) => Math.abs(ours[index]!.x[glyph]! - x)),
		);
		expect(Math.max(...errors)).toBeLessThanOrEqual(3);
		// Word compresses an overfull line so that it ends 0.7 pt past the margin (x = 525.15).
		const ends = word
			.filter((line) => line.font.includes('Arial') && line.x.at(-1)! > 505)
			.map((line) => [line, ours[word.indexOf(line)]!] as const);
		expect(ends.length).toBeGreaterThan(40);
		for (const [theirs, mine] of ends)
			expect(Math.abs(mine.x.at(-1)! - theirs.x.at(-1)!)).toBeLessThanOrEqual(1.5);
	});
});
