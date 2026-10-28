import { describe, expect, it } from 'vitest';
import { resolveScriptFontSet, segmentByFont } from './script-fonts';
import {
	getSubstituteFontFamily,
	hasDirectSubstitution,
	buildFontFamilyString,
} from './font-substitution';

describe('shared font policies', () => {
	it('normalizes lookup while keeping the authored family first', () => {
		expect(getSubstituteFontFamily(' calibri ')).toBe(
			'"calibri", "Carlito", "Liberation Sans", "Arial", sans-serif',
		);
		expect(getSubstituteFontFamily('CALIBRI Bold')).toContain('"Carlito"');
		expect(hasDirectSubstitution('constructor')).toBe(false);
		expect(getSubstituteFontFamily('__proto__')).toBe('"__proto__", sans-serif');
		expect(buildFontFamilyString('A"\\\nB', ['serif'])).toBe('"AB", serif');
	});
	it('resolves inherited faces and PANOSE without any presentation model', () => {
		const fonts = resolveScriptFontSet(
			{ eastAsiaFont: 'SimSun' },
			{
				complexScriptFont: 'Unknown Face',
				complexScriptFontPanose: '02020509000000000000',
			},
			getSubstituteFontFamily('Calibri'),
		);
		expect(fonts.eastAsia).toContain('"SimSun"');
		expect(fonts.complexScript).toContain('monospace');
		expect(fonts.symbol).toBe(fonts.latin);
	});
	it('keeps UTF-16 offsets and combines adjacent equal font choices', () => {
		const text = 'A\u{20000}中عB';
		const pieces = segmentByFont(text, {
			latin: 'Latin',
			eastAsia: 'CJK',
			complexScript: 'Arabic',
			symbol: 'Latin',
		});
		expect(pieces).toEqual([
			{ text: 'A', startIndex: 0, fontFamily: 'Latin' },
			{ text: '\u{20000}中', startIndex: 1, fontFamily: 'CJK' },
			{ text: 'ع', startIndex: 4, fontFamily: 'Arabic' },
			{ text: 'B', startIndex: 5, fontFamily: 'Latin' },
		]);
		expect(
			segmentByFont(text, {
				latin: 'Same',
				eastAsia: 'Same',
				complexScript: 'Same',
				symbol: 'Same',
			}),
		).toEqual([{ text, startIndex: 0, fontFamily: 'Same' }]);
	});
});
