import type { PptxElement, TextSegment } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { fontFamilyStyleUpdates, mergeTextStyleUpdate } from './font-family-update';
import * as coreFontEditing from 'ooxml-core/pptx/editor/render/font-family-update';
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

	it('carries an east Asian face that matched the inherited Latin face', () => {
		const inheriting: TextSegment[] = [
			{ text: 'Sample ', style: { bold: true } },
			{ text: '文字', style: { eastAsiaFont: 'Old Sans', underline: true } },
		];
		const box = textBox(inheriting);
		const expected = [
			{ fontFamily: 'New Serif', bold: true },
			{ fontFamily: 'New Serif', eastAsiaFont: 'New Serif', underline: true },
		];
		const whole = textFontFamilyPatch(box, 'New Serif') as { textSegments: TextSegment[] };
		expect(whole.textSegments.map((run) => run.style)).toStrictEqual(expected);
		const { patch } = applyTextStyleUpdate(box, { fontFamily: 'New Serif' }, null);
		const runs = (patch as { textSegments: TextSegment[] }).textSegments;
		expect(runs.map((run) => run.style)).toStrictEqual(expected);
		const selected = applyTextStyleUpdate(
			box,
			{ fontFamily: 'New Serif' },
			{ startSegIdx: 1, startOffset: 0, endSegIdx: 1, endOffset: 2 },
		).patch as { textSegments: TextSegment[] };
		expect(selected.textSegments[1]?.style).toStrictEqual(expected[1]);
	});
});

it('keeps the UI compatibility exports identical to the core implementation', () => {
	expect(fontFamilyStyleUpdates).toBe(coreFontEditing.fontFamilyStyleUpdates);
	expect(mergeTextStyleUpdate).toBe(coreFontEditing.mergeTextStyleUpdate);
});
