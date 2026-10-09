import { describe, expect, it } from 'vitest';
import { changeTextCase, changeTextCaseRanges } from './change-case';

describe('changeTextCase', () => {
	it('applies every Office mode', () => {
		const text = 'hello WORLD. this is visio!\nnew Line';
		expect(changeTextCase(text, 'upper')).toBe('HELLO WORLD. THIS IS VISIO!\nNEW LINE');
		expect(changeTextCase(text, 'lower')).toBe('hello world. this is visio!\nnew line');
		expect(changeTextCase(text, 'sentence')).toBe('Hello world. This is visio!\nNew line');
		expect(changeTextCase(text, 'capitalize')).toBe('Hello World. This Is Visio!\nNew Line');
		expect(changeTextCase(text, 'toggle')).toBe('HELLO world. THIS IS VISIO!\nNEW lINE');
	});
	it('keeps apostrophes and numbers inside words', () => {
		expect(changeTextCase("don't STOP 3d", 'capitalize')).toBe("Don't Stop 3d");
		expect(changeTextCase('3 apples. ok', 'sentence')).toBe('3 apples. Ok');
		expect(changeTextCase('e.g.x', 'sentence')).toBe('E.g.x');
	});
	it('handles non-BMP and length-changing letters', () => {
		expect(changeTextCase('straße \u{10428}', 'upper')).toBe('STRASSE \u{10400}');
		expect(changeTextCase('漢字', 'toggle')).toBe('漢字');
	});
	it('rejects unknown modes', () => {
		expect(() => changeTextCase('a', 'title' as never)).toThrow(TypeError);
	});
});

describe('changeTextCaseRanges', () => {
	it('returns only changed spans', () => {
		expect(changeTextCaseRanges('ab CD', 'upper')).toEqual([{ start: 0, end: 2, text: 'AB' }]);
		expect(changeTextCaseRanges('AB', 'upper')).toEqual([]);
	});
	it('splits at run boundaries and line breaks while keeping word context', () => {
		expect(changeTextCaseRanges('abcd\nef', 'capitalize', [2])).toEqual([
			{ start: 0, end: 1, text: 'A' },
			{ start: 5, end: 6, text: 'E' },
		]);
		expect(changeTextCaseRanges('abcd\nef', 'upper', [2])).toEqual([
			{ start: 0, end: 2, text: 'AB' },
			{ start: 2, end: 4, text: 'CD' },
			{ start: 5, end: 7, text: 'EF' },
		]);
	});
	it('keeps surrogate pairs whole', () => {
		expect(changeTextCaseRanges('\u{10428}x', 'upper')).toEqual([
			{ start: 0, end: 3, text: '\u{10400}X' },
		]);
	});
});
