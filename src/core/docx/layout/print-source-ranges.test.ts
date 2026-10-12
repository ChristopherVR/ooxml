import { describe, expect, it } from 'vitest';
import { layoutParagraph } from './paragraph-layout';
import { adaptDocumentModel } from './adapter';
import { createDocument } from '../index';
const measurer = { widthOf: (text: string) => text.length * 10, lineHeightOf: () => 20 };
describe('print editing source positions', () => {
	it('maps multi-character atomic labels to one position and keeps following text aligned', () => {
		const layout = layoutParagraph(
			{
				kind: 'paragraph',
				id: 'body',
				runs: [{ text: 'A' }, { text: '12', sourceLength: 1 }, { text: 'B' }],
			},
			200,
			measurer,
			() => {},
		);
		expect(
			layout.lines[0]!.fragments.map((fragment) => [
				fragment.text,
				fragment.sourceStart,
				fragment.sourceEnd,
			]),
		).toEqual([
			['A', 0, 1],
			['12', 1, 2],
			['B', 2, 3],
		]);
	});
	it('counts inline pictures even when they have no source text', () => {
		const layout = layoutParagraph(
			{
				kind: 'paragraph',
				id: 'body',
				runs: [
					{
						text: '',
						object: { partName: 'image', contentType: 'image/png', widthPx: 20, heightPx: 20 },
					},
					{ text: 'B' },
				],
			},
			200,
			measurer,
			() => {},
		);
		expect(
			layout.lines[0]!.fragments.map((fragment) => [fragment.sourceStart, fragment.sourceEnd]),
		).toEqual([
			[0, 1],
			[1, 2],
		]);
	});
	it('adapts note references and equation placeholders with their real editor source lengths', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'body',
				runs: [
					{ text: 'A' },
					{ text: '', noteReference: { kind: 'footnote', id: '2' } },
					{ text: 'B' },
				],
			},
		];
		const paragraph = adaptDocumentModel(model).sections[0]!.blocks[0]!;
		expect(paragraph).toMatchObject({
			runs: [
				{ text: 'A', sourceLength: 1 },
				{ text: '1', sourceLength: 1 },
				{ text: 'B', sourceLength: 1 },
			],
		});
	});
});
