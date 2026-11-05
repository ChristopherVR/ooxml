import { buildEquationBlockStyle } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { getTextBlockStyle } from './element-style';
import { resolveEquationBox } from './equation-box';

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

describe('resolveEquationBox', () => {
	it('lays a pure equation body out with the shared equation block style', () => {
		const el = box([{ text: '', equationXml: { 'm:r': {} }, equationNumber: '1' }]);
		const resolved = resolveEquationBox(el);
		expect(resolved?.style).toStrictEqual(buildEquationBlockStyle(getTextBlockStyle(el)));
		expect(resolved?.equations).toHaveLength(1);
		expect(resolved?.equations[0]?.number).toBe('1');
	});

	it('leaves prose that mixes with a formula on the paragraph path', () => {
		expect(
			resolveEquationBox(box([{ text: 'Area ' }, { text: '', equationXml: { 'm:r': {} } }])),
		).toBeUndefined();
		expect(resolveEquationBox(box([{ text: 'plain' }]))).toBeUndefined();
	});
});
