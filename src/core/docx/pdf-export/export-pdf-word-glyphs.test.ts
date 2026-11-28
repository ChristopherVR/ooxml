// @vitest-environment node
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { exportDocxPdf } from './index';
import { referenceFonts } from './__fixtures__/word-reference/reference';
import { pdfGlyphLines, type PdfGlyphLine } from './__fixtures__/word-sizes/pdf-glyphs';
import reference from './__fixtures__/word-sizes/word-glyphs.json';

/** Word's lines, as `__fixtures__/word-sizes/extract-word.ts` read them from its PDFs. */
type WordLine = Omit<PdfGlyphLine, 'x'> & { x?: number[] };
const WORD: Record<'font-sizes' | 'glyph-positions', WordLine[]> = reference;

async function exported(name: keyof typeof WORD): Promise<PdfGlyphLine[]> {
	const docx = readFileSync(new URL(`./__fixtures__/word-sizes/${name}.docx`, import.meta.url));
	const { model } = await loadDocx(new Uint8Array(docx));
	const { bytes } = await exportDocxPdf(model, {
		fonts: referenceFonts(),
		now: new Date('2026-11-28T00:00:00Z'),
	});
	return pdfGlyphLines(bytes);
}

const rms = (values: number[]) =>
	Math.sqrt(values.reduce((sum, value) => sum + value * value, 0) / values.length);
const worst = (values: number[]) => Math.max(...values.map(Math.abs));

/**
 * The export against Word 16's Save as PDF of the issue #382 documents, glyph by glyph: the sizes
 * written in `Tf`, every baseline and the glyph positions. The tolerances are what the export
 * reaches today; what is left is Word's own advance rule and its line-to-line pitch, open in
 * `docs/parity-checklist.md`. The test fonts (Carlito, Liberation) have the advances and vertical
 * metrics of the fonts Word used (Calibri, Times New Roman, Arial).
 */
describe.each(['font-sizes', 'glyph-positions'] as const)("%s against Word's PDF", async (name) => {
	const ours = await exported(name);
	const word = WORD[name];

	it('draws the same lines, in the same order, on the same pages', () => {
		expect(ours.map(({ page, text }) => ({ page, text }))).toEqual(
			word.map(({ page, text }) => ({ page, text })),
		);
	});

	it("writes every line at Word's drawn size", () => {
		expect(ours.map((line) => Math.round(line.size * 1000) / 1000)).toEqual(
			word.map((line) => line.size),
		);
	});

	it("puts every baseline within 0.25 pt of Word's (about two 1/600 inch steps)", () => {
		const errors = word.map((line, index) => ours[index]!.y - line.y);
		expect(worst(errors)).toBeLessThanOrEqual(0.25);
		expect(rms(errors)).toBeLessThanOrEqual(0.1);
	});

	it("keeps the glyphs near Word's positions", () => {
		const errors = word.flatMap((line, index) =>
			(line.x ?? []).map((x, glyph) => ours[index]!.x[glyph]! - x),
		);
		expect(errors.length).toBeGreaterThan(1000);
		// Short left-aligned lines stay within a third of a point of Word; the long justified lines
		// of glyph-positions.docx drift further, by Word's unknown advances and word spacing.
		const [rmsLimit, maxLimit] = name === 'font-sizes' ? [0.08, 0.36] : [0.2, 0.7];
		expect(rms(errors)).toBeLessThanOrEqual(rmsLimit);
		expect(worst(errors)).toBeLessThanOrEqual(maxLimit);
	});
});
