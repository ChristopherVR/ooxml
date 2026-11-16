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
			// Lengths are compared as numbers: engines round a stored px value differently.
			const px = (value: string): number => Number.parseFloat(value);
			expect(px(first.style.left)).toBeCloseTo(680 / 3, 2);
			const next = [
				...pages[1]!.querySelectorAll<HTMLElement>('.dve-print-custom-footnote-separator'),
			];
			expect(next.map((separator) => separator.textContent)).toEqual(['Continued ', 'Continued ']);
			expect(next.map((separator) => px(separator.style.top))).toEqual([-16, 320]);
			expect(px(next[0]!.style.left)).toBe(0);
			expect(px(next[1]!.style.left)).toBeCloseTo(680 / 3, 2);
			const notes = [...pages[1]!.querySelectorAll<HTMLElement>('.dve-print-footnote')];
			expect(notes.map((note) => note.style.top)).toEqual(['0px', '336px']);
			expect(element.querySelectorAll('.dve-print-footnote-separator')).toHaveLength(0);
		},
	);
});
