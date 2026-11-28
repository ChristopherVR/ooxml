// @vitest-environment node
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { exportDocxPdf } from './index';
import { referenceFonts } from './__fixtures__/word-reference/reference';
import { pdfGlyphLines, type PdfGlyphLine } from './__fixtures__/word-sizes/pdf-glyphs';

/**
 * Line heights of mixed fonts against Word 16's Save as PDF (issues #391 and #393). The reporter's
 * files are not published, so the fixtures are rebuilt from the XML the issues quote
 * (`generate-line-height-fixtures.py`) and checked against the numbers the issues give.
 */
async function firstPage(name: string): Promise<PdfGlyphLine[]> {
	const url = new URL(`../layout/fixtures/word-layout-reports/${name}.docx`, import.meta.url);
	const { model } = await loadDocx(new Uint8Array(readFileSync(url)));
	const { bytes } = await exportDocxPdf(model, {
		fonts: referenceFonts(),
		now: new Date('2026-11-28T00:00:00Z'),
	});
	const lines = await pdfGlyphLines(bytes);
	return lines.filter((line) => line.page === lines[0]!.page);
}

const PAGE_HEIGHT_PT = 841.92;
const pitch = (lines: PdfGlyphLine[]) => (lines[0]!.y - lines[10]!.y) / 10;

describe('line height with two fonts (#391)', () => {
	it("puts the larger ascent, the larger descent and the line gap in the line, as Word's 14.364 pt", async () => {
		const mixed = await firstPage('mixed-font-line-height');
		const control = await firstPage('mixed-font-line-height-control');
		// Times New Roman 10.693 + Garamond descent 3.158 + gap 0.510 = 14.361 pt in Word.
		expect(pitch(mixed)).toBeCloseTo(14.364, 1);
		expect(mixed.length).toBe(48);
		expect(pitch(control)).toBeCloseTo(13.802, 1);
		expect(control.length).toBe(50);
	});
});

describe('a space in a larger run (#393)', () => {
	it('keeps the line at the height of its text, as Word does', async () => {
		const control = await firstPage('larger-space-run-control');
		for (const name of ['larger-space-run', 'larger-space-run-mid-line']) {
			const lines = await firstPage(name);
			expect(lines[0]!.y).toBeCloseTo(control[0]!.y, 2);
			expect(lines.length).toBe(control.length);
			expect(pitch(lines)).toBeCloseTo(pitch(control), 2);
		}
	});
});

describe('exact line spacing (#419)', () => {
	it("puts the baseline at 80 % of the line height, as Word's PDF does", async () => {
		const exact = await firstPage('exact-line-baseline');
		const control = await firstPage('single-spacing-control');
		const topPt = 70.85;
		const heights = [10, 15, 20, 15, 15, 15];
		// Lines stack at their exact heights; the first line's top is the top margin.
		let top = topPt;
		const offsets = exact.map((line, index) => {
			const offset = top + heights[index]! * 0.8;
			top += heights[index]!;
			return [PAGE_HEIGHT_PT - line.y, offset] as const;
		});
		// Word's own baselines are within 0.1 pt of the 80 % rule (8.01, 12.05, 16.01, 11.97).
		for (const [actual, expected] of offsets)
			expect(Math.abs(actual - expected)).toBeLessThan(0.16);
		// Word 16 reads 8.01, 12.05 and 16.01 below the top of the first three lines.
		expect(PAGE_HEIGHT_PT - exact[0]!.y - topPt).toBeCloseTo(8.01, 1);
		expect(exact.length).toBe(6);
		expect(control.length).toBe(6);
	});
});
