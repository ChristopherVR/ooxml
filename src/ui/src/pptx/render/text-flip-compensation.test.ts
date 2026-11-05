import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { getTextCompensationTransform } from './element-style-transform';
import { buildTextBlockStyle } from './text-block-style';
import { getTextCompensationOrigin, getTextCompensationStyle } from './text-flip-compensation';

function shape(overrides: Partial<PptxElement> = {}): PptxElement {
	return {
		id: 's',
		type: 'shape',
		x: 0,
		y: 0,
		width: 400,
		height: 200,
		text: 'Lorem ipsum',
		textStyle: { fontSize: 24 },
		...overrides,
	} as PptxElement;
}

// PowerPoint 16 exports: flipH never mirrors text, flipV turns it upside down
// (rotation + 180), both flips leave the shape's own half turn.
describe('getTextCompensationTransform', () => {
	it('mirrors back exactly when one flip is set, at any rotation', () => {
		expect(getTextCompensationTransform(shape())).toBeUndefined();
		expect(getTextCompensationTransform(shape({ flipHorizontal: true }))).toBe('scaleX(-1)');
		expect(getTextCompensationTransform(shape({ flipVertical: true }))).toBe('scaleX(-1)');
		for (const rotation of [30, 120, 240, 330]) {
			expect(getTextCompensationTransform(shape({ rotation, flipHorizontal: true }))).toBe(
				'scaleX(-1)',
			);
			expect(getTextCompensationTransform(shape({ rotation, flipVertical: true }))).toBe(
				'scaleX(-1)',
			);
		}
		expect(
			getTextCompensationTransform(shape({ flipHorizontal: true, flipVertical: true })),
		).toBeUndefined();
	});
});

describe('getTextCompensationOrigin', () => {
	it('pivots on the box centre for symmetric insets', () => {
		expect(getTextCompensationOrigin(shape({ flipVertical: true }))).toBe('center');
	});

	it('pivots on the text rectangle centre when the insets differ', () => {
		const lopsided = shape({
			flipVertical: true,
			textStyle: { fontSize: 24, bodyInsetLeft: 40, bodyInsetRight: 0 },
		} as Partial<PptxElement>);
		// Text rectangle 40 .. 400: centre 220.
		expect(getTextCompensationOrigin(lopsided)).toBe('220px 50%');
	});

	it('is absent for an unflipped shape', () => {
		expect(getTextCompensationStyle(shape())).toBeUndefined();
	});
});

describe('buildTextBlockStyle body layout', () => {
	it('carries the compensation for the bindings that paint text on one element', () => {
		const style = buildTextBlockStyle(shape({ flipVertical: true }), { bodyLayout: true });
		expect(style.transform).toBe('scaleX(-1)');
		expect(style.transformOrigin).toBe('center');
		expect(buildTextBlockStyle(shape(), { bodyLayout: true }).transform).toBeUndefined();
	});

	it('composes it after the body rotation', () => {
		const rotated = shape({
			flipHorizontal: true,
			textStyle: { fontSize: 24, bodyRotation: 90 },
		} as Partial<PptxElement>);
		const style = buildTextBlockStyle(rotated, { bodyLayout: true });
		expect(String(style.transform)).toMatch(/scaleX\(-1\)$/u);
	});
});
