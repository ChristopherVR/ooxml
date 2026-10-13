import type { PptxElement, TextSegment } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { setFontFamilyPatch } from './editor-text-extra-mutations';

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

	it('carries an east Asian face that matched the inherited Latin face', () => {
		const patch = setFontFamilyPatch(textBox(inheritsLatin), 'New Serif') as {
			textSegments: TextSegment[];
		};

		expect(patch.textSegments[0]?.style).toStrictEqual({
			fontFamily: 'New Serif',
			eastAsiaFont: 'New Serif',
			bold: true,
		});
	});
});
