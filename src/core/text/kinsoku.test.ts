import { describe, expect, it } from 'vitest';
import {
	hangsAtLineEnd,
	kinsokuLanguage,
	kinsokuSets,
	prohibitedAtLineEnd,
	prohibitedAtLineStart,
} from './kinsoku';

describe('kinsokuLanguage', () => {
	it('maps language tags to the four East Asian line-breaking languages', () => {
		expect(kinsokuLanguage('ja-JP')).toBe('ja');
		expect(kinsokuLanguage('zh-CN')).toBe('zh-Hans');
		expect(kinsokuLanguage('zh-SG')).toBe('zh-Hans');
		expect(kinsokuLanguage('zh')).toBe('zh-Hans');
		expect(kinsokuLanguage('zh-TW')).toBe('zh-Hant');
		expect(kinsokuLanguage('zh-HK')).toBe('zh-Hant');
		expect(kinsokuLanguage('zh-Hant-MO')).toBe('zh-Hant');
		expect(kinsokuLanguage('ko-KR')).toBe('ko');
		expect(kinsokuLanguage('en-US')).toBeUndefined();
		expect(kinsokuLanguage(undefined)).toBeUndefined();
	});
});

describe('kinsokuSets', () => {
	it('keeps closing marks off a line start and opening brackets off a line end', () => {
		const japanese = kinsokuSets('ja');
		for (const ch of ['。', '、', '」', '）', '々', '!', '％', 'ゝ'])
			expect(prohibitedAtLineStart(ch, japanese), ch).toBe(true);
		for (const ch of ['「', '（', '$', '￥'])
			expect(prohibitedAtLineEnd(ch, japanese), ch).toBe(true);
		expect(prohibitedAtLineStart('あ', japanese)).toBe(false);
		expect(prohibitedAtLineStart('ぁ', japanese)).toBe(false);
		expect(prohibitedAtLineStart('ー', japanese)).toBe(false);
	});

	it('differs per language as Word does', () => {
		expect(prohibitedAtLineStart('々', kinsokuSets('ja'))).toBe(true);
		expect(prohibitedAtLineStart('々', kinsokuSets('zh-Hans'))).toBe(false);
		expect(prohibitedAtLineStart('～', kinsokuSets('zh-Hans'))).toBe(true);
		expect(prohibitedAtLineStart('～', kinsokuSets('ja'))).toBe(false);
		expect(prohibitedAtLineStart('︰', kinsokuSets('zh-Hant'))).toBe(true);
		expect(prohibitedAtLineEnd('．', kinsokuSets('zh-Hans'))).toBe(true);
		// Korean has no default set.
		expect(kinsokuSets('ko')).toEqual({ noStart: '', noEnd: '' });
	});

	it('adds small kana and the prolonged sound mark to the strict Japanese set only', () => {
		const strict = kinsokuSets('ja', { strict: true });
		for (const ch of ['ぁ', 'ッ', 'ー', 'ｯ', '。'])
			expect(prohibitedAtLineStart(ch, strict), ch).toBe(true);
		expect(kinsokuSets('zh-Hans', { strict: true })).toEqual(kinsokuSets('zh-Hans'));
		expect(strict.noEnd).toBe(kinsokuSets('ja').noEnd);
	});

	it('replaces the set of the custom language only', () => {
		const settings = {
			noLineBreaksBefore: { language: 'ja-JP', chars: '々' },
			noLineBreaksAfter: { language: 'ja-JP', chars: '＄' },
		};
		const japanese = kinsokuSets('ja', settings);
		expect(prohibitedAtLineStart('々', japanese)).toBe(true);
		expect(prohibitedAtLineStart('。', japanese)).toBe(false);
		expect(prohibitedAtLineEnd('（', japanese)).toBe(false);
		expect(prohibitedAtLineEnd('＄', japanese)).toBe(true);
		expect(prohibitedAtLineStart('。', kinsokuSets('zh-Hans', settings))).toBe(true);
	});

	it('tests the base character of a grapheme', () => {
		expect(prohibitedAtLineStart('ゞ', kinsokuSets('ja'))).toBe(true);
		expect(prohibitedAtLineStart('', kinsokuSets('ja'))).toBe(false);
	});
});

describe('hangsAtLineEnd', () => {
	it('hangs the commas, full stops and closing marks Word hangs', () => {
		for (const ch of ['。', '、', '，', '．', '」', '）', '’'])
			expect(hangsAtLineEnd(ch, 'ja'), ch).toBe(true);
		for (const ch of ['！', '？', '・', '：', '〕', ')'])
			expect(hangsAtLineEnd(ch, 'ja'), ch).toBe(false);
		expect(hangsAtLineEnd('〕', 'zh-Hans')).toBe(true);
		expect(hangsAtLineEnd('！', 'zh-Hans')).toBe(true);
		expect(hangsAtLineEnd('．', 'zh-Hans')).toBe(false);
		expect(hangsAtLineEnd('、', 'zh-Hant')).toBe(true);
		expect(hangsAtLineEnd('､', 'zh-Hant')).toBe(false);
		expect(hangsAtLineEnd('。', 'ko')).toBe(false);
	});
});
