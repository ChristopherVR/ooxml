// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { layoutSections } from 'ooxml-core/docx/layout';
import { renderPrintLayout } from './print-layout';
const measurer = { widthOf: (text: string) => text.length * 5, lineHeightOf: () => 20 };
describe('custom footnote separator rendering', () => {
	it('draws styled story text and rules above notes without the default rule', () => {
		const result = layoutSections(
			{
				sections: [
					{
						page: {
							widthPx: 300,
							heightPx: 150,
							marginTopPx: 0,
							marginBottomPx: 0,
							marginLeftPx: 0,
							marginRightPx: 0,
						},
						blocks: [
							{
								kind: 'paragraph',
								id: 'body',
								runs: [{ text: 'Claim' }],
								footnotes: [
									{
										id: '2',
										paragraphs: [{ kind: 'paragraph', id: 'note', runs: [{ text: 'Source' }] }],
									},
								],
							},
						],
						footnoteSeparators: {
							separator: [
								{
									kind: 'paragraph',
									id: 'sep',
									runs: [
										{ text: 'Evidence', bold: true },
										{
											text: '',
											color: '#ff0000',
											object: {
												partName: '',
												contentType: '',
												widthPx: 192,
												heightPx: 1,
												noteSeparator: 'separator',
											},
										},
									],
								},
							],
						},
					},
				],
			},
			measurer,
		);
		const { element } = renderPrintLayout(result);
		const separator = element.querySelector<HTMLElement>('.dve-print-custom-footnote-separator')!;
		expect(separator.textContent).toBe('Evidence');
		expect(separator.style.top).toBe('-20px');
		expect(separator.querySelector<HTMLElement>('.dve-print-note-rule')!.style.width).toBe('192px');
		expect(separator.querySelector<HTMLElement>('.dve-print-note-rule')!.style.color).toBe(
			'rgb(255, 0, 0)',
		);
		expect(element.querySelector('.dve-print-footnote-separator')).toBeNull();
		expect(element.querySelector('.dve-print-footnote')!.textContent).toBe('Source');
	});
});
