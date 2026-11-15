import { describe, expect, it } from 'vitest';
import { complexBreakAllowed, isComplexBreakChar } from './complex-breaks';

const clusters = (text: string) =>
	[...new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(text)].map(
		(part) => part.segment,
	);

describe('isComplexBreakChar', () => {
	it('recognises Thai, Lao, Khmer and Myanmar', () => {
		for (const ch of ['ก', 'ລ', 'ក', 'က', 'ꩠ']) expect(isComplexBreakChar(ch), ch).toBe(true);
		for (const ch of ['a', '中', 'अ', ' ']) expect(isComplexBreakChar(ch), ch).toBe(false);
	});
});

describe('complexBreakAllowed', () => {
	it('allows a break between any two clusters, as Word breaks mid-word', () => {
		const thai = clusters('ระดับเสียง');
		// Word's own break in its probe paragraph: ...ระดั|บเสียง
		expect(complexBreakAllowed('ดั', 'บ')).toBe(true);
		for (let index = 1; index < thai.length; index++) {
			const before = thai[index - 1]!;
			if (before === 'เ') continue;
			expect(complexBreakAllowed(before, thai[index]!), `${before}|${thai[index]}`).toBe(true);
		}
	});

	it('keeps a Thai or Lao leading vowel with its consonant', () => {
		expect(complexBreakAllowed('เ', 'ส')).toBe(false);
		expect(complexBreakAllowed('ไ', 'ท')).toBe(false);
		expect(complexBreakAllowed('ເ', 'ປ')).toBe(false);
		expect(complexBreakAllowed('ย', 'เ')).toBe(true);
	});

	it('keeps stacked Khmer and Myanmar consonants together', () => {
		// Newer segmenters already keep a COENG conjunct in one cluster; older ones end it at the sign.
		expect(complexBreakAllowed('ខ្', 'ម')).toBe(false);
		expect(clusters('ខ្មែរ').at(-1)).toBe('រ');
		expect(complexBreakAllowed('က္', 'က')).toBe(false);
		expect(complexBreakAllowed('ကား', 'သ')).toBe(true);
	});

	it('only applies between two complex-context clusters', () => {
		expect(complexBreakAllowed('ก', 'a')).toBe(false);
		expect(complexBreakAllowed('a', 'ก')).toBe(false);
		expect(complexBreakAllowed('', 'ก')).toBe(false);
	});
});
