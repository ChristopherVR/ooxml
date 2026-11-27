// @vitest-environment node
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { exportDocxPdf } from './index';
import { pdfLines } from './__fixtures__/word-reference/pdf-lines';
import {
	REFERENCES,
	WORD_LINES,
	referenceDocx,
	referenceFonts,
} from './__fixtures__/word-reference/reference';

const now = new Date('2026-11-27T00:00:00Z');

/**
 * Word's "Save as PDF" (Word 16.0.20625, `ExportAsFixedFormat`, with Carlito, Liberation Sans,
 * Liberation Serif and Noto Serif installed) of the reference documents in
 * `__fixtures__/word-reference` (`generate.mjs`), against this export in Node with no measurer
 * and the fonts' Latin subsets supplied: the same pages, and on each page the same lines
 * (`pdfLines`: text per baseline, white space removed). `src/core/scripts/docx-word-pdf-reference.ps1`
 * and `read-word-pdfs.ts` record Word's side in `word-lines.json`.
 */
describe('exportDocxPdf in Node against Word', () => {
	for (const name of REFERENCES)
		it(`${name}: Word's page count and line breaks`, async () => {
			const { model } = await loadDocx(referenceDocx(name));
			const { bytes, report } = await exportDocxPdf(model, { fonts: referenceFonts(), now });
			const word = WORD_LINES[name];
			expect(report.pages).toBe(word.length);
			expect(await pdfLines(bytes)).toEqual(word);
			expect(report.fontMeasurement?.map((entry) => entry.source)).toEqual(['supplied']);
		});

	it('writes the same bytes twice for the same input and `now`', async () => {
		const { model } = await loadDocx(referenceDocx('header-footer'));
		const options = { fonts: referenceFonts(), now, pdfA: true };
		const first = await exportDocxPdf(model, options);
		const second = await exportDocxPdf(model, options);
		expect(Buffer.from(second.bytes).equals(Buffer.from(first.bytes))).toBe(true);
	});

	/** The letter with its font renamed, so no supplied file is named for it. */
	async function letterIn(family: string) {
		const zip = await JSZip.loadAsync(referenceDocx('letter'));
		const styles = await zip.file('word/styles.xml')!.async('string');
		zip.file('word/styles.xml', styles.replaceAll('"Carlito"', `"${family}"`));
		return (await loadDocx(await zip.generateAsync({ type: 'uint8array' }))).model;
	}

	it("measures Calibri with the supplied Carlito, Word's metric-compatible substitute", async () => {
		const { bytes, report } = await exportDocxPdf(await letterIn('Calibri'), {
			fonts: referenceFonts(),
			now,
		});
		expect(report.fontMeasurement).toEqual([
			{ family: 'Calibri', font: 'Carlito', source: 'substitute' },
		]);
		expect(report.substitutedFonts).toEqual([{ family: 'Calibri', standardFont: 'Carlito' }]);
		expect(await pdfLines(bytes)).toEqual(WORD_LINES.letter);
	});

	it("uses the host's font mapping before the substitution table", async () => {
		const { report } = await exportDocxPdf(await letterIn('House Sans'), {
			fonts: referenceFonts(),
			fontSubstitutions: { 'House Sans': 'Carlito' },
			now,
		});
		expect(report.fontMeasurement).toEqual([
			{ family: 'House Sans', font: 'Carlito', source: 'mapping' },
		]);
	});

	it('reports a family it had to estimate because no file was supplied', async () => {
		const { report } = await exportDocxPdf(await letterIn('Unlisted Display'), {
			fonts: referenceFonts(),
			now,
		});
		expect(report.fontMeasurement).toEqual([
			{ family: 'Unlisted Display', font: null, source: 'estimated' },
		]);
		expect(report.notes.join('\n')).toContain(
			'No font file was supplied for Unlisted Display: its text was measured with estimated metrics',
		);
		expect(report.substitutedFonts).toEqual([
			{ family: 'Unlisted Display', standardFont: 'Helvetica' },
		]);
	});
});
