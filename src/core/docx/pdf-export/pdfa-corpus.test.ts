import { mkdirSync, writeFileSync } from 'node:fs';
import { delimiter, join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { buildPdfACorpus, corpusFonts } from './__fixtures__/pdfa-corpus';

// `scripts/pdfa-validate.mjs` sets PDFA_CORPUS_DIR (where to write the files) and PDFA_FONTS
// (font files to embed instead of the synthetic one) and runs veraPDF on the result.
const folder = process.env.PDFA_CORPUS_DIR;
const fontFiles = (process.env.PDFA_FONTS ?? '').split(delimiter).filter(Boolean);

describe('PDF/A corpus', () => {
	it('exports every corpus document as a conformant PDF/A file', async () => {
		const corpus = await buildPdfACorpus(corpusFonts(fontFiles));
		const unembedded = await buildPdfACorpus([]);
		if (folder) {
			mkdirSync(folder, { recursive: true });
			for (const { name, result } of corpus) {
				writeFileSync(join(folder, `${name}.pdf`), result.bytes);
				writeFileSync(join(folder, `${name}.report.json`), JSON.stringify(result.report, null, 2));
			}
			const { name, result } = unembedded[0]!;
			writeFileSync(join(folder, `${name}-standard-fonts.pdf`), result.bytes);
		}
		for (const { name, result } of corpus) {
			expect({ name, pdfA: result.report.pdfA }).toEqual({
				name,
				pdfA: { conformant: true, level: 'U', issues: [] },
			});
			expect(result.report.omitted, name).toEqual([]);
			const raw = Buffer.from(result.bytes).toString('latin1');
			expect(raw).toContain('<pdfaid:conformance>U</pdfaid:conformance>');
			expect(raw, name).not.toMatch(/\/Subtype \/Type1\b/);
		}
		const control = unembedded[0]!.result;
		expect(control.report.pdfA).toEqual({
			conformant: false,
			level: 'B',
			issues: [expect.stringContaining('Calibri')],
		});
		expect(Buffer.from(control.bytes).toString('latin1')).toContain(
			'<pdfaid:conformance>B</pdfaid:conformance>',
		);
	});
});
