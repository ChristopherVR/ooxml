// @vitest-environment jsdom
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadDocx } from 'ooxml-core/docx';
import { layoutDocumentModel } from 'ooxml-core/docx/layout';
import { renderPrintLayout } from './print-layout';
describe('legacy Word note column rendering', () => {
	it.each([12, 14])(
		'mode %s positions each separator immediately above its own bottom-aligned area',
		async (mode) => {
			const loaded = await loadDocx(
				new Uint8Array(
					readFileSync(
						new URL(
							'../../../core/docx/__fixtures__/footnote-columns/mode' +
								mode +
								'-two-custom-eighty.docx',
							import.meta.url,
						),
					),
				),
			);
			const result = layoutDocumentModel(loaded.model, {
				widthOf: (text) => text.length * 5,
				lineHeightOf: () => 16,
			});
			const { element } = renderPrintLayout(result);
			const pages = [...element.querySelectorAll<HTMLElement>('.dve-print-page')];
			expect(pages).toHaveLength(2);
			const first = pages[0]!.querySelector<HTMLElement>('.dve-print-custom-footnote-separator')!;
			expect(first.textContent).toContain('Sources');
			expect(first.style.left).toBe('226.66666666666669px');
			const next = [
				...pages[1]!.querySelectorAll<HTMLElement>('.dve-print-custom-footnote-separator'),
			];
			expect(
				next.map((separator) => [separator.textContent, separator.style.left, separator.style.top]),
			).toEqual([
				['Continued ', '0px', '-16px'],
				['Continued ', '226.66666666666669px', '320px'],
			]);
			const notes = [...pages[1]!.querySelectorAll<HTMLElement>('.dve-print-footnote')];
			expect(notes.map((note) => note.style.top)).toEqual(['0px', '336px']);
			expect(element.querySelectorAll('.dve-print-footnote-separator')).toHaveLength(0);
		},
	);
});
