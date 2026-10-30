import { expect, it } from 'vitest';
import { createDocument, createListDefinition, signedTwips, twips } from '../index';
import { adaptDocumentModel } from './adapter';
import { layoutParagraph } from './paragraph-layout';
import type { TextMeasurer } from './measure';

const measurer: TextMeasurer = { widthOf: (text) => text.length * 10, lineHeightOf: () => 20 };

function laidOut(withPicture: boolean) {
	const model = createDocument();
	const { catalog, numId } = createListDefinition(
		{
			abstractNums: {},
			nums: {},
			pictureBullets: {
				'0': {
					id: '0',
					partName: 'word/media/b.png',
					contentType: 'image/png',
					widthPt: 15,
					heightPt: 15,
				},
			},
			warnings: [],
		},
		[
			{
				level: 0,
				start: 1,
				numFmt: 'bullet',
				lvlText: 'o',
				suffix: 'none',
				indentLeftTwips: signedTwips(720),
				hangingTwips: twips(360),
				...(withPicture ? { pictureBulletId: '0' } : {}),
			},
		],
	);
	model.numberingCatalog = catalog;
	model.blocks = [
		{ type: 'paragraph', id: 'p', numbering: { numId, level: 0 }, runs: [{ text: 'Body' }] },
	];
	const paragraph = adaptDocumentModel(model).sections[0]!.blocks[0]!;
	if (paragraph.kind !== 'paragraph') throw new Error('Missing paragraph');
	return layoutParagraph(paragraph, 500, measurer, () => {});
}

it('measures a picture bullet by the picture width, not the fallback glyph', () => {
	expect(laidOut(false).lines[0]!.fragments[0]).toMatchObject({ text: 'o', widthPx: 10 });
	// 15pt authored width = 20px
	expect(laidOut(true).lines[0]!.fragments[0]).toMatchObject({ widthPx: 20 });
});
