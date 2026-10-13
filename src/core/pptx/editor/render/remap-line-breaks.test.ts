import type { TextSegment } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { remapTextToSegments } from './remap-text';

/**
 * A paragraph whose lines are joined by `a:br`, then a second paragraph, the
 * shape PDF-to-PPTX converters and "Shift+Enter" authors produce. Only the
 * paragraphs carry `a:pPr` spacing.
 */
const SPACED = { lineSpacingExactPt: 12.86, paragraphSpacingAfter: 5 };
function source(): TextSegment[] {
	return [
		{ text: 'First line', style: { bold: true }, paragraphProperties: SPACED },
		{
			text: '\n',
			style: { fontSize: 9 },
			isLineBreak: true,
			breakRunProperties: { '@_sz': '900' },
		},
		{ text: 'Second line', style: {} },
		{ text: '\n', style: {}, isLineBreak: true },
		{ text: 'Third line', style: {} },
		{ text: '\n', style: {}, isParagraphBreak: true },
		{ text: 'Next paragraph', style: {}, paragraphProperties: SPACED },
	];
}

const kinds = (segments: TextSegment[]) =>
	segments
		.filter((segment) => segment.isLineBreak || segment.isParagraphBreak)
		.map((segment) => (segment.isLineBreak ? 'br' : 'p'));

describe('remapTextToSegments with soft line breaks', () => {
	it('keeps an a:br a line break when a character is typed', () => {
		const result = remapTextToSegments(
			'First line\nSecond lineX\nThird line\nNext paragraph',
			source(),
			{},
		);
		// Before: br, br and p all came back as paragraph breaks, so the first
		// paragraph split into three and its lines took paragraph spacing.
		expect(kinds(result)).toEqual(['br', 'br', 'p']);
		expect(result.find((segment) => segment.isLineBreak)).toMatchObject({
			text: '\n',
			style: { fontSize: 9 },
			breakRunProperties: { '@_sz': '900' },
		});
		expect(result.find((segment) => segment.text === 'Second lineX')).toBeDefined();
		expect(result[0].paragraphProperties).toEqual(SPACED);
	});

	it('keeps the remaining lines in one paragraph when a line is deleted', () => {
		const result = remapTextToSegments('First line\nThird line\nNext paragraph', source(), {});
		expect(kinds(result)).toEqual(['br', 'p']);
	});

	it('starts a new paragraph for a line typed with Enter', () => {
		const result = remapTextToSegments(
			'First line\nSecond line\nThird line\nNext paragraph\nTyped',
			source(),
			{},
		);
		expect(kinds(result)).toEqual(['br', 'br', 'p', 'p']);
	});

	it('still separates real paragraphs with paragraph breaks', () => {
		const plain: TextSegment[] = [
			{ text: 'One', style: {} },
			{ text: '\n', style: {} },
			{ text: 'Two', style: {} },
		];
		expect(kinds(remapTextToSegments('OneX\nTwo', plain, {}))).toEqual(['p']);
	});
});
