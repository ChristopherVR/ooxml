// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { layoutSections, type LayoutParagraph } from 'ooxml-core/docx/layout';
import { renderPrintLayout } from './print-layout';

const page = {
	widthPx: 200,
	heightPx: 140,
	marginTopPx: 20,
	marginBottomPx: 20,
	marginLeftPx: 0,
	marginRightPx: 0,
};
const measurer = { widthOf: (text: string) => text.length * 10, lineHeightOf: () => 20 };
const paragraph: LayoutParagraph = {
	kind: 'paragraph',
	id: 'body',
	runs: [{ text: 'Claim' }],
	footnotes: [
		{ id: '1', paragraphs: [{ kind: 'paragraph', id: 'note', runs: [{ text: 'Source' }] }] },
	],
};

describe('Print Layout section footnotes', () => {
	it.each(['center', 'bottom'] as const)(
		'renders %s-aligned body above its footnote area',
		(verticalAlign) => {
			const result = layoutSections(
				{ sections: [{ page, blocks: [paragraph], verticalAlign }] },
				measurer,
			);
			const { element } = renderPrintLayout(result);
			const body = element.querySelector<HTMLElement>('.dve-print-column .dve-print-block')!;
			const notes = element.querySelector<HTMLElement>('.dve-print-footnotes')!;
			expect(body).not.toBeNull();
			expect(notes.textContent).toContain('Source');
			// Block coordinates are relative to the column's top content margin.
			const bottom = page.marginTopPx + parseFloat(body.style.top) + parseFloat(body.style.height);
			expect(bottom).toBeLessThanOrEqual(parseFloat(notes.style.top) - 12);
		},
	);

	it('shows footnotes only on the original page across an odd-page section break', () => {
		const result = layoutSections(
			{
				sections: [
					{ page, blocks: [paragraph] },
					{
						page,
						break: 'oddPage',
						blocks: [{ kind: 'paragraph', id: 'next', runs: [{ text: 'Next' }] }],
					},
				],
			},
			measurer,
		);
		const { element, pageCount } = renderPrintLayout(result);
		expect(pageCount).toBe(3);
		expect(element.querySelectorAll('.dve-print-footnotes')).toHaveLength(1);
		expect(element.textContent?.match(/Source/g)).toHaveLength(1);
	});
});
