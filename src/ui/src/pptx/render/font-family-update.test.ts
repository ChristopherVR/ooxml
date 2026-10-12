import type { PptxElement, TextSegment } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { fontFamilyStyleUpdates, mergeTextStyleUpdate } from './font-family-update';
import { applyStyleToSelectedSegments } from './inline-selection-utils';
import { textFontFamilyPatch, textStylePatch } from './inspector-helpers';
import { applyTextStyleUpdate } from './text-style-update';

/** A run that names one face in every slot, as a CJK deck usually does. */
const sameFace = {
	fontFamily: 'Old Sans',
	eastAsiaFont: 'Old Sans',
	complexScriptFont: 'Old Sans',
};

function textBox(segments: TextSegment[]): PptxElement {
	return {
		type: 'text',
		id: 'box',
		x: 0,
		y: 0,
		width: 100,
		height: 40,
		text: segments.map((segment) => segment.text).join(''),
		textStyle: { ...sameFace },
		textSegments: segments,
	} as PptxElement;
}

describe('fontFamilyStyleUpdates', () => {
	it('carries the east Asian and complex script faces that matched the Latin face', () => {
		expect(fontFamilyStyleUpdates(sameFace, { fontFamily: 'New Serif' })).toStrictEqual({
			fontFamily: 'New Serif',
			eastAsiaFont: 'New Serif',
			complexScriptFont: 'New Serif',
		});
	});

	it('keeps a distinct east Asian face the run chose on purpose', () => {
		const style = { fontFamily: 'Latin Sans', eastAsiaFont: 'Asian Gothic' };
		expect(mergeTextStyleUpdate(style, { fontFamily: 'New Serif' })).toStrictEqual({
			fontFamily: 'New Serif',
			eastAsiaFont: 'Asian Gothic',
		});
	});

	it('follows a theme face that resolved to the Latin face and drops the replaced tokens', () => {
		const style = {
			fontFamily: 'Theme Body',
			latinFontThemeToken: '+mn-lt',
			eastAsiaFont: 'Theme Body',
			eastAsiaFontThemeToken: '+mn-ea',
		};
		expect(mergeTextStyleUpdate(style, { fontFamily: 'New Serif' })).toStrictEqual({
			fontFamily: 'New Serif',
			latinFontThemeToken: undefined,
			eastAsiaFont: 'New Serif',
			eastAsiaFontThemeToken: undefined,
		});
	});

	it('keeps a theme east Asian face that differs from the theme Latin face', () => {
		const style = {
			fontFamily: 'Theme Latin',
			latinFontThemeToken: '+mn-lt',
			eastAsiaFont: 'Theme Gothic',
			eastAsiaFontThemeToken: '+mn-ea',
		};
		expect(mergeTextStyleUpdate(style, { fontFamily: 'New Serif' })).toStrictEqual({
			fontFamily: 'New Serif',
			latinFontThemeToken: undefined,
			eastAsiaFont: 'Theme Gothic',
			eastAsiaFontThemeToken: '+mn-ea',
		});
	});

	it('adds no script faces to a run that only names a Latin face', () => {
		expect(
			fontFamilyStyleUpdates({ fontFamily: 'Old Sans' }, { fontFamily: 'New Serif' }),
		).toStrictEqual({
			fontFamily: 'New Serif',
		});
		expect(fontFamilyStyleUpdates(undefined, { fontFamily: 'New Serif' })).toStrictEqual({
			fontFamily: 'New Serif',
		});
	});

	it('leaves other updates and explicit script faces alone', () => {
		const bold = { bold: true };
		expect(fontFamilyStyleUpdates(sameFace, bold)).toBe(bold);
		expect(
			fontFamilyStyleUpdates(sameFace, { fontFamily: 'New Serif', eastAsiaFont: 'Picked EA' }),
		).toStrictEqual({
			fontFamily: 'New Serif',
			eastAsiaFont: 'Picked EA',
			complexScriptFont: 'New Serif',
		});
	});
});

describe('font family edits on a whole text box', () => {
	const segments: TextSegment[] = [
		{ text: 'Sample ', style: { ...sameFace } },
		{ text: '文字', style: { ...sameFace } },
	];

	it('textFontFamilyPatch reaches every run, not only the body', () => {
		const patch = textFontFamilyPatch(textBox(segments), 'New Serif') as {
			textStyle: Record<string, unknown>;
			textSegments: TextSegment[];
		};
		expect(patch.textStyle.eastAsiaFont).toBe('New Serif');
		for (const segment of patch.textSegments) {
			expect(segment.style.fontFamily).toBe('New Serif');
			expect(segment.style.eastAsiaFont).toBe('New Serif');
		}
	});

	it('textStylePatch alone still only touches the body', () => {
		const patch = textStylePatch(textBox(segments), { fontFamily: 'New Serif' });
		expect('textSegments' in patch).toBe(false);
	});

	it('applyTextStyleUpdate carries the script faces without a selection', () => {
		const { patch } = applyTextStyleUpdate(textBox(segments), { fontFamily: 'New Serif' }, null);
		const runs = (patch as { textSegments: TextSegment[] }).textSegments;
		expect(runs.map((run) => run.style.eastAsiaFont)).toStrictEqual(['New Serif', 'New Serif']);
	});

	it('applies the same rule to a selected range', () => {
		const { newSegments } = applyStyleToSelectedSegments(
			segments,
			{ startSegIdx: 1, startOffset: 0, endSegIdx: 1, endOffset: 2 },
			{ fontFamily: 'New Serif' },
		);
		expect(newSegments[0]?.style.eastAsiaFont).toBe('Old Sans');
		expect(newSegments[1]?.style.eastAsiaFont).toBe('New Serif');
	});
});
