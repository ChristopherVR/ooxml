import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { readFrame } from './morph-text-frame-read';

const style = { fontFamily: 'Calibri', fontSize: 20, color: '#000000' };

function box(
	paragraphs: { text: string; before?: number; after?: number }[],
	textStyle: Record<string, unknown> = {},
): PptxElement {
	const textSegments = paragraphs.flatMap((paragraph, index) => [
		...(index > 0 ? [{ text: '\n', style }] : []),
		{
			text: paragraph.text,
			style,
			paragraphProperties: {
				...(paragraph.before === undefined ? {} : { paragraphSpacingBefore: paragraph.before }),
				...(paragraph.after === undefined ? {} : { paragraphSpacingAfter: paragraph.after }),
			},
		},
	]);
	return {
		id: 'box',
		type: 'text',
		x: 0,
		y: 0,
		width: 400,
		height: 300,
		text: paragraphs.map((paragraph) => paragraph.text).join('\n'),
		textSegments,
		textStyle: {
			fontFamily: 'Calibri',
			fontSize: 20,
			bodyInsetLeft: 0,
			bodyInsetRight: 0,
			bodyInsetTop: 0,
			bodyInsetBottom: 0,
			...textStyle,
		},
		shapeStyle: { fillMode: 'none' },
	} as PptxElement;
}

describe('reading a morphing text frame', () => {
	// PowerPoint 16 exports: the gap between two paragraphs is the first one's
	// space after plus the next one's space before (6 + 12 pt measured 51.75pt
	// between 28pt lines). The flex column adds the two margins, so the frame
	// carries the whole gap below the upper paragraph.
	it('carries the next paragraph space before in the gap below a paragraph', () => {
		const frame = readFrame(
			box([{ text: 'one', after: 8 }, { text: 'two', before: 16 }, { text: 'three' }]),
		)!;
		expect(frame.paragraphs.map((paragraph) => paragraph.spaceAfterPx)).toEqual([24, 0, 0]);
		expect(frame.paragraphs.map((paragraph) => paragraph.indented)).toEqual([false, false, false]);
	});

	it('keeps the first paragraph space before above it only with spcFirstLastPara', () => {
		const lines = [{ text: 'one', before: 10 }, { text: 'two' }];
		expect(readFrame(box(lines))!.paragraphs[0]!.spaceBeforePx).toBeUndefined();
		const first = readFrame(box(lines, { spaceFirstLastParagraph: true }))!.paragraphs[0]!;
		expect(first.spaceBeforePx).toBe(10);
		expect(first.indented).toBe(true);
	});

	it('flags a distributed or justified vertical anchor', () => {
		expect(readFrame(box([{ text: 'one' }]))!.anchorSpread).toBeUndefined();
		expect(readFrame(box([{ text: 'one' }], { vAlign: 'distributed' }))!.anchorSpread).toBe(true);
	});
});
