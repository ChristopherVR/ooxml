// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { layoutSections, type LayoutParagraph } from 'ooxml-core/docx/layout';
import { renderPrintLayout } from './print-layout';
import css from './styles/print.css?raw';

const page = {
	widthPx: 200,
	heightPx: 100,
	marginTopPx: 0,
	marginBottomPx: 0,
	marginLeftPx: 0,
	marginRightPx: 0,
};
const measurer = { widthOf: (text: string) => text.length * 10, lineHeightOf: () => 20 };

describe('continued footnotes in Print Layout', () => {
	it('renders every line once, with stable note identity and full-width continuation separators', () => {
		const block: LayoutParagraph = {
			kind: 'paragraph',
			id: 'ref',
			runs: [{ text: 'Claim' }],
			footnotes: [
				{
					id: '2',
					paragraphs: [
						{
							kind: 'paragraph',
							id: 'note',
							runs: [
								{
									text: Array.from({ length: 10 }, (_, index) => `Line${index}`).join('\n'),
									italic: true,
								},
							],
						},
					],
				},
			],
		};
		const result = layoutSections({ sections: [{ page, blocks: [block] }] }, measurer);
		const { element, pageCount } = renderPrintLayout(result);
		expect(pageCount).toBe(3);
		expect(
			[...element.querySelectorAll('.dve-print-footnote')].map(
				(note) => (note as HTMLElement).dataset.noteId,
			),
		).toEqual(['2', '2', '2']);
		expect(element.querySelectorAll('[data-continued="true"]')).toHaveLength(2);
		expect(element.querySelectorAll('.dve-print-footnote-continuation-separator')).toHaveLength(2);
		expect(css).toMatch(/\.dve-print-footnote-continuation-separator\s*\{\s*width:\s*100%/);
		expect(
			[...element.querySelectorAll('.dve-print-footnote .dve-print-line')].map(
				(line) => line.textContent,
			),
		).toEqual(Array.from({ length: 10 }, (_, index) => `Line${index}`));
		for (const span of element.querySelectorAll<HTMLElement>('.dve-print-footnote span'))
			expect(span.style.fontStyle).toBe('italic');
	});
});
