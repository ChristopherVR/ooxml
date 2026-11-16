import { describe, expect, it } from 'vitest';

import { isSymbolEncodedFont, symbolCharToUnicode, symbolTextToUnicode } from './symbol-fonts';

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
		expect(symbolCharToUnicode('Wingdings', 'K')).toBeUndefined();
		expect(symbolCharToUnicode('Arial', '§')).toBeUndefined();
		expect(symbolTextToUnicode('Wingdings', 'nK')).toBeUndefined();
		expect(symbolTextToUnicode('Wingdings 2', 'n')).toBeUndefined();
		expect(symbolTextToUnicode('Wingdings', 'n n')).toBe('■ ■');
	});

	it('substitutes one UTF-16 unit per glyph so run offsets stay valid', () => {
		for (const family of ['Symbol', 'Wingdings']) {
			for (let code = 0x20; code <= 0xff; code++) {
				const mapped = symbolCharToUnicode(family, String.fromCharCode(code));
				if (mapped !== undefined) expect(mapped.length, `${family} ${code}`).toBe(1);
			}
		}
	});
});
