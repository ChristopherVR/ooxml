import type { TextSegment } from 'pptx-viewer-core';
import { describe, it, expect } from 'vitest';

import { resolveParagraphStrutFontSize } from './paragraph-strut';

function segment(fontSize: number | undefined): Pick<TextSegment, 'style' | 'bulletInfo' | 'text'> {
	return { style: { fontSize }, text: 'x' } as Pick<TextSegment, 'style' | 'bulletInfo' | 'text'>;
}

describe('resolveParagraphStrutFontSize', () => {
	it('returns the SMALLEST run size when it differs from the body default', () => {
		// Not the largest: a bigger run already sizes its own line correctly
		// through ordinary inline layout (see the module doc comment), so using
		// it here would inflate every OTHER wrapped line of the paragraph too.
		expect(resolveParagraphStrutFontSize([segment(12), segment(20)], 16)).toBe(12);
	});

	it('returns undefined when the paragraph matches the body default', () => {
		expect(resolveParagraphStrutFontSize([segment(16)], 16)).toBeUndefined();
	});

	it('scales the returned strut size by fontScale (autofit shrink)', () => {
		// Authored 40px run, autofit shrinks every run to 80% (32px): the strut
		// must re-base to the SHRUNK size, or the line box stays sized for the
		// unshrunk text while the run itself renders smaller.
		expect(resolveParagraphStrutFontSize([segment(40)], 16, 0.8)).toBe(32);
	});

	it('defaults fontScale to 1 (no autofit)', () => {
		expect(resolveParagraphStrutFontSize([segment(40)], 16)).toBe(40);
	});

	it('ignores bullet segments when finding the smallest run', () => {
		const bulletSeg = {
			style: { fontSize: 1 },
			bulletInfo: { char: '•' },
			text: '• ',
		} as Pick<TextSegment, 'style' | 'bulletInfo' | 'text'>;
		expect(resolveParagraphStrutFontSize([bulletSeg, segment(14)], 16)).toBe(14);
	});

	it('counts the first run of an a:buNone paragraph', () => {
		const noBullet = {
			style: { fontSize: 12 },
			bulletInfo: { none: true, ownedByParagraph: true },
			text: 'body',
		} as Pick<TextSegment, 'style' | 'bulletInfo' | 'text'>;
		expect(resolveParagraphStrutFontSize([noBullet], 17)).toBe(12);
	});

	it('counts a text run that carries a bullet set in the editor', () => {
		// List commands put `bulletInfo` on the paragraph's first text run.
		const bulleted = {
			style: { fontSize: 12 },
			bulletInfo: { char: '•' },
			text: 'body',
		} as Pick<TextSegment, 'style' | 'bulletInfo' | 'text'>;
		expect(resolveParagraphStrutFontSize([bulleted], 17)).toBe(12);
	});

	// Regression for audit-text/gen.py slide 16 ("small HUGE small wraps onto
	// the / second line of text"): a 48pt run must not push a wrapped 12pt
	// line down by inflating the whole paragraph's strut to 48pt.
	it('does not let one oversized run in the paragraph inflate the strut', () => {
		expect(resolveParagraphStrutFontSize([segment(12), segment(48), segment(12)], 16)).toBe(12);
	});
});
