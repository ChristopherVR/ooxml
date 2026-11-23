import { describe, expect, it } from 'vitest';
import {
	hyphenationPoints,
	isAllCapsWord,
	SOFT_HYPHEN,
	softHyphenIndices,
	stripSoftHyphens,
} from './hyphenation';

function splitWord(word: string, options?: Parameters<typeof hyphenationPoints>[1]): string {
	const points = hyphenationPoints(word, options);
	if (!points.length) return word;
	let out = '';
	let prev = 0;
	for (const p of points) {
		out += `${word.slice(prev, p)}-`;
		prev = p;
	}
	return out + word.slice(prev);
}

describe('hyphenationPoints', () => {
	it('splits multi-syllable English words at syllable boundaries', () => {
		expect(hyphenationPoints('internationalization').length).toBeGreaterThanOrEqual(4);
		expect(splitWord('hyphenation')).toBe('hy-phen-a-tion');
		expect(splitWord('paragraph')).toBe('pa-ra-graph');
		expect(splitWord('formatting')).toBe('for-mat-ting');
		expect(splitWord('running')).toBe('run-ning');
		expect(splitWord('computer')).toBe('com-pu-ter');
		expect(splitWord('table')).toBe('ta-ble');
	});

	it('preserves leading and trailing punctuation offsets', () => {
		expect(splitWord('"formatting,"')).toBe('"for-mat-ting,"');
		expect(splitWord('(paragraph)')).toBe('(pa-ra-graph)');
	});

	it('does not hyphenate short words, digits, URLs, or non-Latin text', () => {
		expect(hyphenationPoints('cat')).toEqual([]);
		expect(hyphenationPoints('word')).toEqual([]);
		expect(hyphenationPoints('make')).toEqual([]);
		expect(hyphenationPoints('123456')).toEqual([]);
		expect(hyphenationPoints('https://example.com')).toEqual([]);
		expect(hyphenationPoints('foo_bar_baz')).toEqual([]);
	});

	it('respects doNotHyphenateCaps for ALL-CAPS words while still splitting title-case words', () => {
		expect(isAllCapsWord('INTERNATIONAL')).toBe(true);
		expect(isAllCapsWord('International')).toBe(false);
		expect(hyphenationPoints('INTERNATIONAL', { doNotHyphenateCaps: true })).toEqual([]);
		expect(
			hyphenationPoints('INTERNATIONAL', { doNotHyphenateCaps: false }).length,
		).toBeGreaterThan(0);
		expect(hyphenationPoints('International', { doNotHyphenateCaps: true }).length).toBeGreaterThan(
			0,
		);
	});
});

describe('soft hyphen helpers', () => {
	it('finds internal soft hyphen indices and strips soft hyphens', () => {
		const word = `co${SOFT_HYPHEN}op${SOFT_HYPHEN}erate`;
		expect(softHyphenIndices(word)).toEqual([2, 5]);
		expect(stripSoftHyphens(word)).toBe('cooperate');
		expect(softHyphenIndices(`${SOFT_HYPHEN}lead`)).toEqual([]);
		expect(softHyphenIndices(`trail${SOFT_HYPHEN}`)).toEqual([]);
		expect(stripSoftHyphens('plain')).toBe('plain');
	});
});
