import type { PptxElement, TextSegment } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { setFontFamilyPatch } from './editor-text-extra-mutations';

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

describe('setFontFamilyPatch on a whole text box', () => {
	it('reaches every run, not only the body style', () => {
		const patch = setFontFamilyPatch(textBox(), 'New Serif') as {
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
