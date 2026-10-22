import { describe, expect, it } from 'vitest';
import { applyFieldTextCase, fieldTokens } from './field-instruction';

describe('field instruction words', () => {
	it('reads the field type, quoted and bare arguments and switches', () => {
		expect(fieldTokens(' STYLEREF "Doc Number" \\l \\* MERGEFORMAT ')).toEqual([
			{ kind: 'text', value: 'STYLEREF' },
			{ kind: 'text', value: 'Doc Number' },
			{ kind: 'switch', value: 'l' },
			{ kind: 'switch', value: '*' },
			{ kind: 'text', value: 'MERGEFORMAT' },
		]);
	});

	it('keeps escaped quotes and backslashes inside quotes, and reads a switch without a space', () => {
		expect(fieldTokens('REF "Say \\"hi\\" at C:\\\\docs" \\*Upper \\H')).toEqual([
			{ kind: 'text', value: 'REF' },
			{ kind: 'text', value: 'Say "hi" at C:\\docs' },
			{ kind: 'switch', value: '*' },
			{ kind: 'text', value: 'Upper' },
			{ kind: 'switch', value: 'h' },
		]);
	});

	it('reads a backslash followed by more than two letters as a word', () => {
		expect(fieldTokens('STYLEREF \\abc')).toEqual([
			{ kind: 'text', value: 'STYLEREF' },
			{ kind: 'text', value: '\\abc' },
		]);
	});
});

describe('field text case formats', () => {
	it('changes every letter for Upper and Lower', () => {
		expect(applyFieldTextCase('Doc Number 4', 'upper')).toBe('DOC NUMBER 4');
		expect(applyFieldTextCase('Doc Number 4', 'lower')).toBe('doc number 4');
	});

	it('capitalizes the first word for FirstCap and every word for Caps', () => {
		expect(applyFieldTextCase('the (draft) plan', 'firstcap')).toBe('The (draft) plan');
		expect(applyFieldTextCase('the (draft) plan', 'caps')).toBe('The (Draft) Plan');
		expect(applyFieldTextCase('éclair au café', 'caps')).toBe('Éclair Au Café');
	});
});
