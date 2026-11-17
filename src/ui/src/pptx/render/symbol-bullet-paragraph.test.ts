import { expect, it } from 'vitest';
import type { PptxElement } from 'ooxml-core/pptx';
import { buildParagraphs } from './text-paragraphs';

it('maps a marker separately from Georgia content without changing the saved model (#290)', () => {
	const element = {
		id: 'symbol-bullet',
		type: 'text',
		x: 0,
		y: 0,
		width: 400,
		height: 100,
		textStyle: { fontFamily: 'Arial', fontSize: 24 },
		textSegments: [
			{
				text: '§',
				style: { fontFamily: 'Wingdings' },
				bulletInfo: { char: '§', fontFamily: 'Wingdings' },
			},
			{ text: 'Georgia paragraph', style: { fontFamily: 'Georgia', fontSize: 24 } },
		],
	} as PptxElement;
	const before = structuredClone(element);
	const paragraph = buildParagraphs(element)[0];
	expect(paragraph?.bulletMarker).toBe('▪');
	expect(paragraph?.bulletStyle.fontFamily).toMatch(/^"Georgia",.*"OOXML Noto Symbols"/u);
	expect(paragraph?.bulletStyle.fontFamily).not.toContain('Wingdings');
	expect(paragraph?.runs.map((run) => run.text).join('')).toBe('Georgia paragraph');
	expect(element).toEqual(before);
});
