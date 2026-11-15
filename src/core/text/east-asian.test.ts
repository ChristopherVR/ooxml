import { describe, expect, it } from 'vitest';
import { eastAsianAutoSpace, isEastAsianChar, isEastAsianLetter } from './east-asian';

describe('isEastAsianChar', () => {
	it('classifies CJK ideographs, kana, Hangul and fullwidth forms as East Asian', () => {
		for (const ch of ['中', 'あ', 'ア', '가', '。', 'Ａ'])
			expect(isEastAsianChar(ch), ch).toBe(true);
		for (const ch of ['a', ' ', '1', 'é']) expect(isEastAsianChar(ch), ch).toBe(false);
	});
});

describe('eastAsianAutoSpace', () => {
	it('spaces East Asian letters from Western letters and digits, in either order', () => {
		expect(eastAsianAutoSpace('本', 'W')).toBe('alpha');
		expect(eastAsianAutoSpace('a', '日')).toBe('alpha');
		expect(eastAsianAutoSpace('本', '1')).toBe('digit');
		expect(eastAsianAutoSpace('2', 'あ')).toBe('digit');
		expect(eastAsianAutoSpace('가', 'W')).toBe('alpha');
		expect(eastAsianAutoSpace('ア', 'ж')).toBe('alpha');
		expect(eastAsianAutoSpace('日', 'é')).toBe('alpha');
	});

	it('leaves punctuation, spaces, fullwidth letters and Western pairs unspaced', () => {
		expect(eastAsianAutoSpace('日', '.')).toBeUndefined();
		expect(eastAsianAutoSpace('(', '日')).toBeUndefined();
		expect(eastAsianAutoSpace('。', 'a')).toBeUndefined();
		expect(eastAsianAutoSpace('日', ' ')).toBeUndefined();
		expect(eastAsianAutoSpace('日', 'Ａ')).toBeUndefined();
		expect(eastAsianAutoSpace('a', 'b')).toBeUndefined();
		expect(isEastAsianLetter('。')).toBe(false);
	});
});
