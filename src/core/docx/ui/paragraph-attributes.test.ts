import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import type { Paragraph } from '../model';
import { loadDocx } from '../index';
import { paragraphAttrs, paragraphFromAttrs } from './paragraph-attributes';

describe('shared paragraph attribute conversion', () => {
	it.each(
		['alignment', 'spacing', 'indent', 'multiple'].flatMap((name) =>
			['before', 'tracked', 'rejected'].map((state) => `${name}-${state}`),
		),
	)('round-trips native %s properties and history without a DOM or global schema', async (name) => {
		const loaded = await loadDocx(
			new Uint8Array(
				await readFile(
					new URL(`../__fixtures__/review-paragraph-formatting/${name}.docx`, import.meta.url),
				),
			),
		);
		const paragraph = loaded.model.blocks[0]!;
		if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
		expect(paragraphFromAttrs(paragraphAttrs(paragraph), paragraph.id, paragraph.runs)).toEqual(
			paragraph,
		);
	});
	it('clones revision metadata while retaining the supplied identity and runs', () => {
		const paragraph: Paragraph = {
			type: 'paragraph',
			id: 'stable',
			runs: [{ text: 'Text' }],
			align: 'left',
			keepNext: false,
			formatRevision: {
				kind: 'paragraphChange',
				id: 'format',
				author: 'Ada',
				previousParagraphPropertiesXml:
					'<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
			},
		};
		const converted = paragraphFromAttrs(paragraphAttrs(paragraph), paragraph.id, paragraph.runs);
		expect(converted).toEqual(paragraph);
		expect(converted.runs).toBe(paragraph.runs);
		converted.formatRevision!.author = 'Changed';
		expect(paragraph.formatRevision!.author).toBe('Ada');
	});
	it('uses the existing branded-unit conversion at the attribute boundary', () => {
		const paragraph = paragraphFromAttrs(
			{
				spacingBeforeTwips: -1,
				spacingAfterTwips: 12.6,
				indentLeftTwips: -12.6,
				lineSpacingTwips: Number.NaN,
			},
			'units',
			[],
		);
		expect(paragraph.spacingBeforeTwips).toBeUndefined();
		expect(paragraph.spacingAfterTwips).toBe(13);
		expect(paragraph.indentLeftTwips).toBe(-13);
		expect(paragraph.lineSpacingTwips).toBeUndefined();
	});
});
