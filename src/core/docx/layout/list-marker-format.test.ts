import { expect, it } from 'vitest';
import { createDocument, createListDefinition, twips } from '../index';
import { adaptDocumentModel } from './adapter';

it('formats the Print Layout marker run from the level, not the paragraph text', () => {
	const model = createDocument();
	const { catalog, numId } = createListDefinition(undefined, [
		{
			level: 0,
			start: 1,
			numFmt: 'decimal',
			lvlText: '%1.',
			markerFormat: { fontFamily: 'Georgia', fontSizeHalfPoints: 28, bold: true, color: '#1f4e79' },
		},
	]);
	model.numberingCatalog = catalog;
	model.blocks = [
		{ type: 'paragraph', id: 'p', numbering: { numId, level: 0 }, runs: [{ text: 'Body' }] },
	];
	const paragraph = adaptDocumentModel(model).sections[0]!.blocks[0]!;
	if (paragraph.kind !== 'paragraph') throw new Error('Missing paragraph');
	const marker = paragraph.runs.find((run) => run.synthetic)!;
	expect(marker).toMatchObject({
		fontFamily: 'Georgia',
		fontSizePt: 14,
		bold: true,
		color: '#1f4e79',
	});
});

it('adds the level tab stop to the paragraph tab stops in order', () => {
	const model = createDocument();
	const { catalog, numId } = createListDefinition(undefined, [
		{ level: 0, start: 1, numFmt: 'decimal', lvlText: '%1.', tabStopTwips: twips(1440) },
	]);
	model.numberingCatalog = catalog;
	model.blocks = [
		{
			type: 'paragraph',
			id: 'p',
			numbering: { numId, level: 0 },
			tabStops: [{ posTwips: twips(2880), align: 'left' }],
			runs: [{ text: 'Body' }],
		} as never,
	];
	const paragraph = adaptDocumentModel(model).sections[0]!.blocks[0]!;
	if (paragraph.kind !== 'paragraph') throw new Error('Missing paragraph');
	expect(paragraph.tabStops?.map((stop) => stop.posPx)).toEqual([96, 192]);
});

it.each([
	[{}, 10],
	[{ fontSizeHalfPoints: 28 }, 14],
])(
	'formats an empty list item from the paragraph mark and level properties %j',
	(markerFormat, size) => {
		const model = createDocument();
		const { catalog, numId } = createListDefinition(undefined, [
			{ level: 0, start: 1, numFmt: 'decimal', lvlText: '%1.', markerFormat },
		]);
		model.numberingCatalog = catalog;
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p',
				numbering: { numId, level: 0 },
				runs: [{ text: '', fontSize: 18 }],
				sourceParagraphPropertiesXml:
					'<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:sz w:val="20"/><w:b/></w:rPr></w:pPr>',
			},
		];
		const paragraph = adaptDocumentModel(model).sections[0]?.blocks[0];
		if (paragraph?.kind !== 'paragraph') throw new Error('Missing paragraph');
		expect(paragraph.runs.find((run) => run.synthetic)).toMatchObject({
			fontSizePt: size,
			bold: true,
		});
	},
);
