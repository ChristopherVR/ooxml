import { describe, expect, it } from 'vitest';
import { autoCorrectTextInput } from './autocorrect';

describe('Word AutoCorrect and AutoFormat As You Type', () => {
	it('converts straight quotes into opening and closing smart quotes', () => {
		expect(autoCorrectTextInput('', '"')).toEqual({ kind: 'replace', deleteBefore: 0, text: '“' });
		expect(autoCorrectTextInput('Hello ', '"')).toEqual({
			kind: 'replace',
			deleteBefore: 0,
			text: '“',
		});
		expect(autoCorrectTextInput('“Hello', '"')).toEqual({
			kind: 'replace',
			deleteBefore: 0,
			text: '”',
		});
		expect(autoCorrectTextInput('', "'")).toEqual({ kind: 'replace', deleteBefore: 0, text: '‘' });
		expect(autoCorrectTextInput('don', "'")).toEqual({
			kind: 'replace',
			deleteBefore: 0,
			text: '’',
		});
	});

	it('triggers automatic bulleted and numbered lists at paragraph start', () => {
		expect(autoCorrectTextInput('-', ' ')).toEqual({
			kind: 'list',
			listKind: 'bullet',
			deleteBefore: 1,
		});
		expect(autoCorrectTextInput('*', ' ')).toEqual({
			kind: 'list',
			listKind: 'bullet',
			deleteBefore: 1,
		});
		expect(autoCorrectTextInput('1.', ' ')).toEqual({
			kind: 'list',
			listKind: 'decimal',
			deleteBefore: 2,
		});
		expect(autoCorrectTextInput('1)', ' ')).toEqual({
			kind: 'list',
			listKind: 'decimal',
			deleteBefore: 2,
		});
		// Mid-paragraph `- ` is not a list trigger
		expect(autoCorrectTextInput('hello -', ' ')).toEqual({
			kind: 'replace',
			deleteBefore: 1,
			text: '– ',
		});
	});

	it('replaces dashes, ellipsis, symbols, fractions and common typos preserving case', () => {
		expect(autoCorrectTextInput('wait-', '-')).toEqual({
			kind: 'replace',
			deleteBefore: 1,
			text: '—',
		});
		expect(autoCorrectTextInput('wait..', '.')).toEqual({
			kind: 'replace',
			deleteBefore: 2,
			text: '…',
		});
		expect(autoCorrectTextInput('(c', ')')).toEqual({
			kind: 'replace',
			deleteBefore: 2,
			text: '©',
		});
		expect(autoCorrectTextInput('(tm', ')')).toEqual({
			kind: 'replace',
			deleteBefore: 3,
			text: '™',
		});
		expect(autoCorrectTextInput('1/2', ' ')).toEqual({
			kind: 'replace',
			deleteBefore: 3,
			text: '½ ',
		});
		expect(autoCorrectTextInput('Teh', ' ')).toEqual({
			kind: 'replace',
			deleteBefore: 3,
			text: 'The ',
		});
		expect(autoCorrectTextInput('TEH', ' ')).toEqual({
			kind: 'replace',
			deleteBefore: 3,
			text: 'THE ',
		});
	});

	it('respects individual option flags', () => {
		expect(autoCorrectTextInput('', '"', { smartQuotes: false })).toBeUndefined();
		expect(autoCorrectTextInput('-', ' ', { autoLists: false })).toBeUndefined();
		expect(autoCorrectTextInput('1/2', ' ', { fractions: false })).toBeUndefined();
		expect(autoCorrectTextInput('teh', ' ', { commonTypos: false })).toBeUndefined();
	});
});
