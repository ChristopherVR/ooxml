import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import { buildTextBodyLayoutStyle } from './text-body-layout';
import { shrinkBlockOffset } from './text-shrink-block';

// A monospace stand-in for the canvas: every character is half the font size wide.
vi.mock('./morph-text-frame-read', async (importOriginal) => ({
	...(await importOriginal<typeof import('./morph-text-frame-read')>()),
	measureMorphTextWithCanvas: (text: string, font: string) => {
		const [, size = '0'] = /([\d.]+)px/u.exec(font) ?? [];
		return [...text].length * Number(size) * 0.5;
	},
}));

function box(text: string, textStyle: Record<string, unknown> = {}, width = 150): PptxElement {
	const style = { fontFamily: 'Calibri', fontSize: 20 };
	return {
		id: 'b',
		type: 'text',
		x: 0,
		y: 0,
		width,
		height: 200,
		text,
		textSegments: [{ text, style: { ...style, color: '#000000' } }],
		textStyle: {
			...style,
			anchorCenter: true,
			bodyInsetLeft: 0,
			bodyInsetRight: 0,
			bodyInsetTop: 0,
			bodyInsetBottom: 0,
			...textStyle,
		},
		shapeStyle: { fillMode: 'none' },
	} as PptxElement;
}

// PowerPoint 16 exports of wrapped `anchorCtr` boxes: the block shrinks to its
// longest line and that box is centred in the text area, whatever the alignment.
describe('shrink-wrapped anchorCtr block', () => {
	it('moves a wrapped left aligned block so its longest line is centred', () => {
		// Lines of 140 and 20 px in a 150 px column: the block is 5 px right of the column.
		expect(shrinkBlockOffset(box('aaaa bbbb cccc dd'))).toBe(5);
		expect(buildTextBodyLayoutStyle(box('aaaa bbbb cccc dd')).marginLeft).toBe('5px');
	});

	it('moves a right aligned block left and leaves a centred one in place', () => {
		expect(shrinkBlockOffset(box('aaaa bbbb cccc dd', { align: 'right' }))).toBe(-5);
		expect(shrinkBlockOffset(box('aaaa bbbb cccc dd', { align: 'center' }))).toBe(0);
		expect(
			buildTextBodyLayoutStyle(box('aaaa bbbb cccc dd', { align: 'center' })).marginLeft,
		).toBeUndefined();
	});

	it('centres the longest line by the text area, inside the insets', () => {
		// 20 px insets leave a 110 px area centred at 75: the lines "aaaa bbbb"
		// (90 px) and "cccc" span 20..110, centred at 65, so the block moves 10 px.
		const inset = box('aaaa bbbb cccc', { bodyInsetLeft: 20, bodyInsetRight: 20 });
		expect(shrinkBlockOffset(inset)).toBe(10);
	});

	it('keeps fit-content where the lines cannot be measured', () => {
		expect(shrinkBlockOffset(box('aaaa bbbb', { columnCount: 2 }))).toBeUndefined();
	});
});
