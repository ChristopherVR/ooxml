import { describe, expect, it } from 'vitest';

import {
	isSymbolEncodedFont,
	officeSymbolFontGlyphs,
	symbolCharToUnicode,
	symbolTextToUnicode,
} from './symbol-fonts';

describe('symbol-encoded fonts', () => {
	it('recognises the Office symbol fonts by family, ignoring case, quotes and fallbacks', () => {
		expect(isSymbolEncodedFont('Wingdings')).toBe(true);
		expect(isSymbolEncodedFont('"wingdings 2", sans-serif')).toBe(true);
		expect(isSymbolEncodedFont('Symbol')).toBe(true);
		expect(isSymbolEncodedFont('Webdings')).toBe(true);
		expect(isSymbolEncodedFont('Georgia')).toBe(false);
		expect(isSymbolEncodedFont(undefined)).toBe(false);
	});

	it("maps the PowerPoint bullet library's Wingdings glyphs (issue #290)", () => {
		expect(symbolCharToUnicode('Wingdings', '§')).toBe('▪');
		expect(symbolCharToUnicode('Wingdings', 'q')).toBe('❑');
		expect(symbolCharToUnicode('Wingdings', 'v')).toBe('❖');
		expect(symbolCharToUnicode('Wingdings', 'Ø')).toBe('➢');
		expect(symbolCharToUnicode('Wingdings', 'ü')).toBe('✔');
		expect(symbolCharToUnicode('Wingdings', 'n')).toBe('■');
		expect(symbolCharToUnicode('Wingdings', 'l')).toBe('●');
	});

	it('reads private-use code points as their byte', () => {
		expect(symbolCharToUnicode('Wingdings', '')).toBe('▪');
		expect(symbolCharToUnicode('Symbol', '')).toBe('•');
	});

	it('maps Symbol Greek letters and operators', () => {
		expect(symbolTextToUnicode('Symbol', 'abg')).toBe('αβγ');
		expect(symbolTextToUnicode('Symbol', 'W')).toBe('Ω');
		expect(symbolTextToUnicode('Symbol', '£³¹')).toBe('≤≥≠');
	});

	it('maps the circled-digit ranges', () => {
		expect(symbolCharToUnicode('Wingdings', '\u0081')).toBe('①');
		expect(symbolCharToUnicode('Wingdings', '\u008a')).toBe('⑩');
		expect(symbolCharToUnicode('Wingdings', '\u008c')).toBe('❶');
	});

	it('returns undefined for unknown glyphs and ordinary fonts', () => {
		expect(symbolCharToUnicode('Wingdings', '\u007f')).toBeUndefined();
		expect(symbolCharToUnicode('Arial', '§')).toBeUndefined();
		expect(symbolTextToUnicode('Wingdings', 'n\u007f')).toBeUndefined();
		expect(symbolTextToUnicode('Wingdings', 'n n')).toBe('■ ■');
	});

	it('maps Webdings and Wingdings 2 and 3 from their byte or private-use code', () => {
		expect(symbolCharToUnicode('Webdings', '!')).toBe('\u{1F577}');
		expect(symbolCharToUnicode('Wingdings 2', '\uF021')).toBe('\u{1F58A}');
		expect(symbolCharToUnicode('Wingdings 3', '!')).toBe('\u2B60');
		expect(officeSymbolFontGlyphs('Wingdings 2')).toContain('\u{1F58A}');
		expect(officeSymbolFontGlyphs('Arial')).toEqual([]);
	});
});
