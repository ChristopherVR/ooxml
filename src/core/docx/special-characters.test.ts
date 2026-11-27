import { describe, expect, it } from 'vitest';
import { symbolCharacter } from './special-characters';

describe('symbolCharacter', () => {
	it('reads a code below U+0100 in a symbol font as its private-use character', () => {
		expect(symbolCharacter('B7', 'Symbol')).toBe('');
		expect(symbolCharacter('00FC', 'Wingdings')).toBe('');
		expect(symbolCharacter('F0FC', 'Wingdings')).toBe('');
	});

	it('reads a code below U+0100 in a text font as the Unicode character (issue #365)', () => {
		// Word 16 shows Calibri 00A7 and 00B6 as the section and pilcrow signs.
		expect(symbolCharacter('00A7', 'Calibri')).toBe('§');
		expect(symbolCharacter('00B6', 'Calibri')).toBe('¶');
		expect(symbolCharacter('2022', 'Calibri')).toBe('•');
		// Segoe UI Symbol is a Unicode font despite its name.
		expect(symbolCharacter('00A7', 'Segoe UI Symbol')).toBe('§');
	});

	it('rejects codes that are not characters', () => {
		expect(symbolCharacter('D800', 'Symbol')).toBe('');
		expect(symbolCharacter('zz', 'Symbol')).toBe('');
	});
});
