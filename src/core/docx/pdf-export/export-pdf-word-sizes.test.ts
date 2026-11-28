// @vitest-environment node
import { readFileSync } from 'node:fs';
import { inflateSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { exportDocxPdf } from './index';
import { referenceFonts } from './__fixtures__/word-reference/reference';

/** The decoded content streams of a PDF (fonts and images left out). */
function content(bytes: Uint8Array): string {
	const raw = Buffer.from(bytes).toString('latin1');
	const streams: string[] = [];
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		if (/\/Filter \/FlateDecode/.test(match[1]!) && !/\/Subtype|\/Length1/.test(match[1]!))
			streams.push(inflateSync(bytes.subarray(start, start + length)).toString('latin1'));
	}
	return streams.join('\n');
}

/** Whether `pt` is a whole number of 1/600 inch (0.12 pt). */
const onGrid = (pt: number) => Math.abs(pt / 0.12 - Math.round(pt / 0.12)) < 1e-3;

/**
 * `font-sizes.docx` (issue #382): one line per size from 8 to 72 pt by half points, in Calibri,
 * Times New Roman and Arial. Word 16 draws every run at the size rounded to 1/600 inch.
 */
describe("exportDocxPdf on Word's 1/600 inch grid", async () => {
	const docx = readFileSync(new URL('./__fixtures__/word-sizes/font-sizes.docx', import.meta.url));
	const { model } = await loadDocx(new Uint8Array(docx));
	const { bytes } = await exportDocxPdf(model, {
		fonts: referenceFonts(),
		now: new Date('2026-11-28T00:00:00Z'),
	});
	const raw = Buffer.from(bytes).toString('latin1');
	const ops = content(bytes);
	// The content is drawn in CSS pixels under a 0.75 scale.
	const sizes = [...ops.matchAll(/\/[EF]\d+ ([\d.]+) Tf/g)].map((match) => Number(match[1]) * 0.75);

	it('writes every font size in whole 1/600 inch, as Word does', () => {
		expect(sizes.length).toBeGreaterThan(100);
		expect(sizes.filter((size) => !onGrid(size))).toEqual([]);
		// 8, 10, 10.5, 11, 13, 13.5 and 50 pt as Word draws them.
		for (const drawn of [8.04, 9.96, 10.56, 11.04, 12.96, 13.56, 50.04])
			expect(sizes.some((size) => Math.abs(size - drawn) < 1e-4)).toBe(true);
		expect(sizes.some((size) => Math.abs(size - 11) < 1e-4)).toBe(false);
	});

	it('rounds the page box and puts baselines on whole 1/600 inch', () => {
		const boxes = [...raw.matchAll(/\/MediaBox \[[^\]]*\]/g)].map(([box]) => box);
		expect(boxes.length).toBeGreaterThan(0);
		// A4 (11906 x 16838 twips), as Word writes it.
		expect(new Set(boxes)).toEqual(new Set(['/MediaBox [0 0 595.32 841.92]']));
		const baselines = [...ops.matchAll(/ -1 [-\d.]+ ([-\d.]+) Tm/g)].map(
			(m) => Number(m[1]) * 0.75,
		);
		expect(baselines.length).toBeGreaterThan(100);
		expect(baselines.filter((y) => !onGrid(y))).toEqual([]);
	});

	it('corrects the glyphs drawn at the rounded size with TJ offsets', () => {
		expect(ops).toMatch(/\[<[\da-f]+> -?\d+ <[\da-f]+>[^\]]*\] TJ/);
	});
});
