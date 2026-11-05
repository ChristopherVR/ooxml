/**
 * PowerPoint never mirrors text: a horizontally flipped shape paints its text
 * as the unflipped one does, a vertically flipped one paints it upside down
 * (PowerPoint 16 exports). The container style flips the whole element, so
 * this binding's text block has to mirror its text back; it painted flipped
 * text mirrored before the shared text block style carried the compensation.
 */
import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { getTextBlockStyle } from './element-style';

function textShape(overrides: Partial<PptxElement>): PptxElement {
	return {
		id: 't1',
		type: 'shape',
		x: 0,
		y: 0,
		width: 200,
		height: 100,
		text: 'hello',
		textStyle: {},
		...overrides,
	} as PptxElement;
}

describe('vue flipped text', () => {
	it('mirrors the text of a shape flipped one way back', () => {
		expect(getTextBlockStyle(textShape({ flipHorizontal: true })).transform).toBe('scaleX(-1)');
		// flipV: the container's scaleY(-1) and this scaleX(-1) make a half turn.
		expect(getTextBlockStyle(textShape({ flipVertical: true })).transform).toBe('scaleX(-1)');
	});

	it('leaves an unflipped shape and a doubly flipped one (a half turn) alone', () => {
		expect(getTextBlockStyle(textShape({})).transform).toBeUndefined();
		const both = textShape({ flipHorizontal: true, flipVertical: true });
		expect(getTextBlockStyle(both).transform).toBeUndefined();
	});

	it('turns about the text rectangle when the insets differ', () => {
		const style = getTextBlockStyle(
			textShape({
				flipVertical: true,
				textStyle: { bodyInsetLeft: 40, bodyInsetRight: 0 },
			} as Partial<PptxElement>),
		);
		expect(style.transformOrigin).toBe('120px 50%');
	});
});
