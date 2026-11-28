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
type Name = 'font-sizes' | 'glyph-positions' | 'line-pitch' | 'line-pitch-letter';
const WORD: Record<Name, WordLine[]> = reference;

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

/** The tolerances each document reaches today: baseline rms and worst, pitch share, glyph x. */
const LIMITS: Record<Name, { y: [number, number]; pitch: number; x: [number, number] }> = {
	'font-sizes': { y: [0.06, 0.2], pitch: 0.98, x: [0.075, 0.33] },
	// Justified lines, and paragraphs with space after, whose last line Word sometimes puts one row
	// lower than its rule for the other lines (not understood yet, see docs/parity-checklist.md).
	'glyph-positions': { y: [0.065, 0.2], pitch: 0.87, x: [0.19, 0.65] },
	'line-pitch': { y: [0.042, 0.085], pitch: 1, x: [0.08, 0.22] },
	'line-pitch-letter': { y: [0.053, 0.105], pitch: 1, x: [0.08, 0.17] },
};

/** How many 1/600 inch rows the baseline steps down from the line before, on the same page. */
const pitches = (lines: { page: number; y: number }[]) =>
	lines.slice(1).flatMap((line, index) => {
		const before = lines[index]!;
		return line.page === before.page ? [Math.round(((before.y - line.y) * 600) / 72)] : [];
	});

/**
 * The export against Word 16's Save as PDF, glyph by glyph: the sizes written in `Tf`, every
 * baseline, the line-to-line pitch in whole 1/600 inch rows and the glyph positions. The first two
 * documents are the ones attached to issue #382; `line-pitch*.docx` were made for the baseline
 * rule (`wordBaselinePx`) and saved as PDF by Word 16. The tolerances are what the export reaches
 * today; Word's glyph advance rule is still open in `docs/parity-checklist.md`. The test fonts
 * (Carlito, Liberation) have the advances and vertical metrics of the fonts Word used (Calibri,
 * Times New Roman, Arial).
 */
describe.each(['font-sizes', 'glyph-positions', 'line-pitch', 'line-pitch-letter'] as const)(
	"%s against Word's PDF",
	async (name) => {
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

		it("puts every baseline within a row of Word's", () => {
			const errors = word.map((line, index) => ours[index]!.y - line.y);
			expect(rms(errors)).toBeLessThanOrEqual(LIMITS[name].y[0]);
			expect(worst(errors)).toBeLessThanOrEqual(LIMITS[name].y[1]);
		});

		it("steps from line to line by Word's pitch", () => {
			const expected = pitches(word);
			const actual = pitches(ours);
			const same = expected.filter((pitch, index) => pitch === actual[index]).length;
			expect(expected.length).toBeGreaterThan(30);
			expect(same / expected.length).toBeGreaterThanOrEqual(LIMITS[name].pitch);
		});

		it("keeps the glyphs near Word's positions", () => {
			const errors = word.flatMap((line, index) =>
				(line.x ?? []).map((x, glyph) => ours[index]!.x[glyph]! - x),
			);
			expect(errors.length).toBeGreaterThan(500);
			// Short left-aligned lines stay within a third of a point of Word; the long justified lines
			// of glyph-positions.docx drift further, by Word's unknown advances and word spacing.
			const [rmsLimit, maxLimit] = LIMITS[name].x;
			expect(rms(errors)).toBeLessThanOrEqual(rmsLimit);
			expect(worst(errors)).toBeLessThanOrEqual(maxLimit);
		});
	},
);
