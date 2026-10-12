import type { PptxElement, TextSegment } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { setFontFamily } from './editor-format-mutations';

const sameFace = { fontFamily: 'Old Sans', eastAsiaFont: 'Old Sans' };

function textBox(): PptxElement {
	return {
		type: 'text',
		id: 'box',
		x: 0,
		y: 0,
		width: 100,
		height: 40,
		text: 'Sample',
		textStyle: { ...sameFace },
		textSegments: [{ text: 'Sample', style: { ...sameFace } }],
	} as PptxElement;
}

describe('setFontFamily on a whole text box', () => {
	it('carries the east Asian face along with the Latin one', () => {
		const patch = setFontFamily(textBox(), 'New Serif') as {
			textStyle: Record<string, unknown>;
			textSegments: TextSegment[];
		};

		expect(patch.textStyle).toMatchObject({ fontFamily: 'New Serif', eastAsiaFont: 'New Serif' });
		expect(patch.textSegments[0]?.style).toMatchObject({
			fontFamily: 'New Serif',
			eastAsiaFont: 'New Serif',
		});
	});
});
