import { buildEquationBlockStyle } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { renderEquationBox } from './EquationBox';

const box = (segments: unknown[]): PptxElement =>
	({
		id: 'e1',
		type: 'text',
		x: 0,
		y: 0,
		width: 300,
		height: 100,
		textStyle: {},
		textSegments: segments,
	}) as unknown as PptxElement;

describe('renderEquationBox', () => {
	it('paints a pure equation body with the shared equation block style', () => {
		const node = renderEquationBox(
			box([{ text: '', equationXml: { 'm:r': { 'm:t': 'x' } } }]),
			false,
		);
		const style = node?.props.style as Record<string, unknown>;
		const expected = buildEquationBlockStyle({}) as Record<string, unknown>;
		expect(style.alignItems).toBe(expected.alignItems);
		expect(style.flexWrap).toBe('nowrap');
		expect(style.fontFamily).toBe(expected.fontFamily);
	});

	it('returns null for prose or an empty body', () => {
		expect(
			renderEquationBox(box([{ text: 'Area ' }, { text: '', equationXml: {} }]), false),
		).toBeNull();
		expect(renderEquationBox(box([{ text: 'plain' }]), false)).toBeNull();
	});
});
