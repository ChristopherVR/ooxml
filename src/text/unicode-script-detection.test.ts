import { describe, expect, it } from 'vitest';
import {
	detectFontScript,
	segmentByScript,
	resolveFontForScript,
	hasDistinctScriptFonts,
} from './index.js';

describe('script font primitives', () => {
	it.each([
		['AéΩЖ', 'latin'],
		['中あア한', 'eastAsia'],
		['אعकก', 'complexScript'],
		['→∑■😀', 'symbol'],
	] as const)('classifies %s into the existing %s font group', (text, category) => {
		for (const character of text)
			expect(detectFontScript(character.codePointAt(0)!)).toBe(category);
	});

	it('preserves supplementary characters and UTF-16 offsets across mixed scripts', () => {
		const text = 'A𠀀😀عZ';
		const runs = segmentByScript(text);
		expect(runs).toEqual([
			{ text: 'A', script: 'latin', startIndex: 0 },
			{ text: '𠀀', script: 'eastAsia', startIndex: 1 },
			{ text: '😀', script: 'symbol', startIndex: 3 },
			{ text: 'ع', script: 'complexScript', startIndex: 5 },
			{ text: 'Z', script: 'latin', startIndex: 6 },
		]);
		expect(runs.map((run) => run.text).join('')).toBe(text);
		for (const run of runs)
			expect(text.slice(run.startIndex, run.startIndex + run.text.length)).toBe(run.text);
	});

	it('keeps ASCII punctuation with the preceding script and initial neutrals with latin', () => {
		expect(segmentByScript(' (中, ع! A)')).toEqual([
			{ text: ' (', script: 'latin', startIndex: 0 },
			{ text: '中, ', script: 'eastAsia', startIndex: 2 },
			{ text: 'ع! ', script: 'complexScript', startIndex: 5 },
			{ text: 'A)', script: 'latin', startIndex: 8 },
		]);
	});

	it('does not lose authored line breaks, empty input or isolated surrogate code units', () => {
		expect(segmentByScript('')).toEqual([]);
		const text = '中\n\t中\ud800x\udc00';
		expect(
			segmentByScript(text)
				.map((run) => run.text)
				.join(''),
		).toBe(text);
		expect(segmentByScript(text)[0]).toEqual({
			text: '中\n\t中',
			script: 'eastAsia',
			startIndex: 0,
		});
	});

	it('uses supplied script families with the historical empty-string fallback to latin', () => {
		const fonts = { latin: 'Base', eastAsia: 'CJK', complexScript: '', symbol: 'Symbols' };
		expect(resolveFontForScript('latin', fonts)).toBe('Base');
		expect(resolveFontForScript('eastAsia', fonts)).toBe('CJK');
		expect(resolveFontForScript('complexScript', fonts)).toBe('Base');
		expect(resolveFontForScript('symbol', fonts)).toBe('Symbols');
		expect(resolveFontForScript('eastAsia', {})).toBeUndefined();
	});

	it('avoids a split without a latin base and when all supplied script fonts match it', () => {
		expect(hasDistinctScriptFonts({ eastAsia: 'CJK' })).toBe(false);
		expect(hasDistinctScriptFonts({ latin: 'Base', eastAsia: 'Base', complexScript: '' })).toBe(
			false,
		);
		expect(hasDistinctScriptFonts({ latin: 'Base', symbol: 'Symbols' })).toBe(true);
	});
});
