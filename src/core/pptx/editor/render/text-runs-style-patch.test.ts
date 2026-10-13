import { describe, expect, it } from 'vitest';
import type { PptxElement, TextPptxElement } from '../../core/types';
import { textRunsStylePatch } from './text-runs-style-patch';
import { updateTextSegmentStyle } from './update-text-segment-style';

const box = (): TextPptxElement => ({
	type: 'text',
	id: 'box',
	x: 0,
	y: 0,
	width: 100,
	height: 40,
	text: 'A文字',
	textStyle: { fontFamily: 'Old Sans', eastAsiaFont: 'Old Sans', latinFontThemeToken: '+mn-lt' },
	textSegments: [
		{ text: 'A', style: { bold: true } },
		{
			text: '文字',
			style: {
				eastAsiaFont: 'Old Sans',
				complexScriptFont: 'Distinct Arabic',
				symbolFont: 'Wingdings',
			},
		},
		{
			text: '',
			style: {},
			isParagraphBreak: true,
			paragraphInsertionStyle: { fontFamily: 'Old Sans', eastAsiaFont: 'Old Sans', italic: true },
		},
	],
});

describe('whole-text style patch', () => {
	it('uses pre-edit body inheritance, updates future paragraph text and preserves other faces', () => {
		const source = box();
		const before = structuredClone(source);
		const patch = textRunsStylePatch(source, { fontFamily: 'Georgia' });
		expect(source).toEqual(before);
		expect(patch.textStyle).toMatchObject({
			fontFamily: 'Georgia',
			eastAsiaFont: 'Georgia',
			latinFontThemeToken: undefined,
		});
		expect(patch.textSegments?.[0]?.style).toEqual({ bold: true, fontFamily: 'Georgia' });
		expect(patch.textSegments?.[1]?.style).toEqual({
			fontFamily: 'Georgia',
			eastAsiaFont: 'Georgia',
			complexScriptFont: 'Distinct Arabic',
			symbolFont: 'Wingdings',
		});
		expect(patch.textSegments?.[2]?.paragraphInsertionStyle).toEqual({
			fontFamily: 'Georgia',
			eastAsiaFont: 'Georgia',
			italic: true,
		});
	});
	it('keeps insertion-only edits distinct from the existing run body', () => {
		const segment = box().textSegments![2]!;
		const edited = updateTextSegmentStyle(
			segment,
			{ fontFamily: 'Georgia' },
			{ updateBodyStyle: false },
		);
		expect(edited.style).toBe(segment.style);
		expect(edited.paragraphInsertionStyle?.fontFamily).toBe('Georgia');
		expect(segment.paragraphInsertionStyle?.fontFamily).toBe('Old Sans');
	});
	it('does not manufacture runs and refuses elements without text properties', () => {
		const source = box();
		delete source.textSegments;
		expect(textRunsStylePatch(source, { fontFamily: 'Georgia' })).not.toHaveProperty(
			'textSegments',
		);
		expect(
			textRunsStylePatch({ type: 'image', id: 'image' } as PptxElement, { bold: true }),
		).toEqual({});
	});
});
