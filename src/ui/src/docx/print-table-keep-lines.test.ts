// @vitest-environment jsdom
import { URL as NodeURL } from 'node:url';
import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { loadDocx } from 'ooxml-core/docx';
import { layoutDocumentModel } from 'ooxml-core/docx/layout';
import { renderPrintLayout } from './print-layout';
const root = new NodeURL('../../../core/docx/__fixtures__/table-keep-lines/', import.meta.url);
const evidence = JSON.parse(
	readFileSync(new NodeURL('native-measurements.json', root), 'utf8'),
) as {
	cases: { name: string; measurements: { lines: { text: string }[]; borders: unknown[] }[] }[];
};
for (const native of evidence.cases)
	it('paints native kept row bands and source ink once: ' + native.name, async () => {
		const loaded = await loadDocx(
			new Uint8Array(readFileSync(new NodeURL(native.name + '.docx', root))),
		);
		const { element } = renderPrintLayout(
			layoutDocumentModel(loaded.model, {
				widthOf: (text) => text.length * 5,
				lineHeightOf: () => 16,
			}),
		);
		const pages = [...element.querySelectorAll('.dve-print-page')];
		expect(pages).toHaveLength(native.measurements.length);
		for (const [index, page] of pages.entries()) {
			const sample = native.measurements[index]!;
			expect(page.querySelectorAll('.dve-print-table')).toHaveLength(sample.borders.length ? 1 : 0);
			expect((page.textContent?.match(/T[01]R\d+C\d+L[A-Z]\d{2}/g) ?? []).sort()).toEqual(
				sample.lines
					.map((line) => line.text.trim())
					.filter((text) => /^T[01]/.test(text))
					.sort(),
			);
			expect(page.querySelectorAll('.dve-print-nested-table')).toHaveLength(0);
			for (const cell of page.querySelectorAll<HTMLElement>('.dve-print-cell'))
				expect(parseFloat(cell.style.borderTopWidth)).toBe(4);
		}
	});
