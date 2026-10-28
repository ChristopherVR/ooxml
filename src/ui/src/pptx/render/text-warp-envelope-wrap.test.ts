import { describe, expect, it } from 'vitest';

import { DEFAULT_BODY_INSET_LR_PX } from '../constants';
import type { EnvelopeSegmentInput } from './text-warp-envelope-types';
import { envelopeWrapWidth, wrapEnvelopeParagraph } from './text-warp-envelope-wrap';
import type { EnvelopeWrapItem } from './text-warp-envelope-wrap';

const ADVANCE = 10;

function items(text: string, segmentIndex = 0): EnvelopeWrapItem[] {
	const seg: EnvelopeSegmentInput = { text, font: { fontSizePx: 20 }, segmentIndex };
	return [...text].map((char) => ({ seg, char, advance: ADVANCE }));
}

function rowTexts(rows: EnvelopeWrapItem[][]): string[] {
	return rows.map((row) => row.map((item) => item.char).join(''));
}

describe('envelopeWrapWidth', () => {
	it('is the box width less the OOXML default insets when the body sets none', () => {
		expect(envelopeWrapWidth(200, {})).toBeCloseTo(200 - 2 * DEFAULT_BODY_INSET_LR_PX, 6);
	});

	it('subtracts the authored left and right insets', () => {
		expect(envelopeWrapWidth(200, { bodyInsetLeft: 20, bodyInsetRight: 30 })).toBe(150);
		expect(envelopeWrapWidth(200, { bodyInsetLeft: 0, bodyInsetRight: 0 })).toBe(200);
	});

	it('wraps for wrap="square" and never for wrap="none"', () => {
		expect(
			envelopeWrapWidth(200, { textWrap: 'square', bodyInsetLeft: 0, bodyInsetRight: 0 }),
		).toBe(200);
		expect(envelopeWrapWidth(200, { textWrap: 'none' })).toBeUndefined();
	});

	it('does not wrap without a text body, and never goes negative', () => {
		expect(envelopeWrapWidth(200, undefined)).toBeUndefined();
		expect(envelopeWrapWidth(10, { bodyInsetLeft: 8, bodyInsetRight: 8 })).toBe(0);
	});
});

describe('wrapEnvelopeParagraph', () => {
	it('keeps the paragraph one row with no wrap width', () => {
		const rows = wrapEnvelopeParagraph(items('one two three four'), undefined);
		expect(rowTexts(rows)).toStrictEqual(['one two three four']);
	});

	it('breaks a paragraph wider than the width into rows at word boundaries', () => {
		// 'one two' is 70 wide; adding ' three' would make 130.
		const rows = wrapEnvelopeParagraph(items('one two three four five'), 100);
		expect(rowTexts(rows)).toStrictEqual(['one two', 'three four', 'five']);
	});

	it('fills a row exactly to the width', () => {
		expect(rowTexts(wrapEnvelopeParagraph(items('ab cd ef'), 50))).toStrictEqual(['ab cd', 'ef']);
	});

	it('narrower widths (larger insets) give more rows', () => {
		const text = 'one two three four five';
		const wide = wrapEnvelopeParagraph(items(text), envelopeWrapWidth(300, {}));
		const narrow = wrapEnvelopeParagraph(
			items(text),
			envelopeWrapWidth(300, { bodyInsetLeft: 100, bodyInsetRight: 100 }),
		);
		expect(wide).toHaveLength(1);
		expect(narrow.length).toBeGreaterThan(1);
	});

	it('never breaks at a no-break space', () => {
		// 'aa bb' at width 30: the ordinary space breaks, the no-break spaces hold.
		expect(rowTexts(wrapEnvelopeParagraph(items('aa bb'), 30))).toEqual(['aa', 'bb']);
		for (const nbsp of ['\u00A0', '\u2007', '\u202F']) {
			expect(rowTexts(wrapEnvelopeParagraph(items(`aa${nbsp}bb`), 30))).toEqual([`aa${nbsp}bb`]);
		}
	});

	it('keeps a single word wider than the width whole on one row', () => {
		expect(rowTexts(wrapEnvelopeParagraph(items('schedule'), 25))).toStrictEqual(['schedule']);
		expect(rowTexts(wrapEnvelopeParagraph(items('a schedule b'), 25))).toStrictEqual([
			'a',
			'schedule',
			'b',
		]);
	});

	it('drops the whitespace at a soft break but keeps leading and trailing whitespace', () => {
		expect(rowTexts(wrapEnvelopeParagraph(items(' aaa   bbb '), 50))).toStrictEqual([
			' aaa',
			'bbb ',
		]);
	});

	it('starts a new row at a hard line break, even when the text fits', () => {
		const paragraph = [...items('ab'), ...items('\n', 1), ...items('cd', 2)];
		const rows = wrapEnvelopeParagraph(paragraph, 1000);
		expect(rowTexts(rows)).toStrictEqual(['ab', 'cd']);
		expect(rows[1][0].seg.segmentIndex).toBe(2);
	});

	it('wraps across runs, carrying each character with its own run', () => {
		const paragraph = [...items('one tw', 0), ...items('o three', 1)];
		const rows = wrapEnvelopeParagraph(paragraph, 70);
		expect(rowTexts(rows)).toStrictEqual(['one two', 'three']);
		expect(rows[0].map((item) => item.seg.segmentIndex)).toStrictEqual([0, 0, 0, 0, 0, 0, 1]);
	});

	it('returns one empty row for an empty paragraph', () => {
		expect(wrapEnvelopeParagraph([], 100)).toStrictEqual([[]]);
	});
});
