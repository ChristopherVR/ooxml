import type { PptxElement, TextSegment } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { setFontFamily } from './editor-format-mutations';

const sameFace = { fontFamily: 'Old Sans', eastAsiaFont: 'Old Sans' };

/** A run that inherits its Latin face from the box but names the old face in `a:ea`. */
const inheritsLatin = { eastAsiaFont: 'Old Sans', bold: true };

function textBox(runStyle: TextSegment['style'] = sameFace): PptxElement {
	return {
		type: 'text',
		id: 'box',
		x: 0,
		y: 0,
		width: 100,
		height: 40,
		text: 'Sample',
		textStyle: { ...sameFace },
		textSegments: [{ text: 'Sample', style: { ...runStyle } }],
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

	it('carries an east Asian face that matched the inherited Latin face', () => {
		const patch = setFontFamily(textBox(inheritsLatin), 'New Serif') as {
			textSegments: TextSegment[];
		};

		expect(patch.textSegments[0]?.style).toStrictEqual({
			fontFamily: 'New Serif',
			eastAsiaFont: 'New Serif',
			bold: true,
		});
	});
});
