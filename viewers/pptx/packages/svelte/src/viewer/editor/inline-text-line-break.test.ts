/**
 * Regression: committing an inline edit turned every `a:br` of the edited body
 * into a paragraph break, so the paragraph's later lines came back as new
 * paragraphs without its `a:pPr` and spacing appeared between its lines.
 */
import type { PptxElement, TextSegment } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { remapInlineText } from './inline-text';

function element(): PptxElement {
	return {
		id: 'el-1',
		type: 'text',
		x: 0,
		y: 0,
		width: 200,
		height: 80,
		text: 'First line\nSecond line\nNext paragraph',
		textSegments: [
			{ text: 'First line', style: {}, paragraphProperties: { paragraphSpacingAfter: 5 } },
			{ text: '\n', style: {}, isLineBreak: true },
			{ text: 'Second line', style: {} },
			{ text: '\n', style: {}, isParagraphBreak: true },
			{ text: 'Next paragraph', style: {} },
		],
	} as unknown as PptxElement;
}

const breaks = (segments: TextSegment[]) =>
	segments.flatMap((s) => (s.isLineBreak ? ['br'] : s.isParagraphBreak ? ['p'] : []));

describe('remapInlineText with a soft line break', () => {
	it('keeps the a:br inside its paragraph when a character is typed', () => {
		const { textSegments } = remapInlineText(element(), 'First lineX\nSecond line\nNext paragraph');
		expect(breaks(textSegments)).toEqual(['br', 'p']);
	});
});
