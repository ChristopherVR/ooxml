import { describe, expect, it } from 'vitest';
import { compileFindPattern, findMatches } from './find-pattern';
import { FindPatternError } from './find-syntax';
import { replacementParts } from './find-replacement';
import { wildcardSource } from './find-wildcards';

const found = (text: string, query: string, options = {}) =>
	findMatches(text, query, options).map((match) => text.slice(match.start, match.end));

const wild = (text: string, query: string) => found(text, query, { wildcards: true });

describe('findMatches: plain search', () => {
	it('matches case-insensitively unless Match case is on', () => {
		expect(found('Cat cat CAT', 'cat')).toEqual(['Cat', 'cat', 'CAT']);
		expect(found('Cat cat CAT', 'cat', { matchCase: true })).toEqual(['cat']);
	});

	it('escapes regular-expression syntax in the query', () => {
		expect(found('a.b a*b a+b', 'a.b')).toEqual(['a.b']);
		expect(found('f(x) f[x]', 'f[x]')).toEqual(['f[x]']);
		expect(found('a\\b', 'a\\b')).toEqual(['a\\b']);
	});

	it('finds whole words only with Find whole words only', () => {
		expect(found('the there other the,', 'the', { wholeWords: true })).toEqual(['the', 'the']);
	});

	it('finds prefixes and suffixes with Match prefix and Match suffix', () => {
		expect(found('pre prefix repre', 'pre', { matchPrefix: true })).toEqual(['pre', 'pre']);
		expect(found('work homework workers', 'work', { matchSuffix: true })).toEqual(['work', 'work']);
	});

	it('skips punctuation inside a match with Ignore punctuation characters', () => {
		expect(found('Mr. Smith, Mr Smith MrSmith', 'mr smith', { ignorePunctuation: true })).toEqual([
			'Mr. Smith',
			'Mr Smith',
		]);
	});

	it('skips white space inside a match with Ignore white-space characters', () => {
		expect(found('a b  ab a b', 'a b', { ignoreWhitespace: true })).toEqual(['a b', 'ab', 'a b']);
	});

	it('does not match a paragraph mark, line break or object', () => {
		expect(found('foo\rbar', 'foobar')).toEqual([]);
		expect(found('fo\vobar', 'foobar')).toEqual([]);
		expect(found(`foo\u0000bar`, 'foobar')).toEqual([]);
	});

	it('skips soft hyphens inside a match and never matches a typed one', () => {
		expect(found('Hyphen­ation', 'hyphenation')).toEqual(['Hyphen­ation']);
		expect(found('­hyphen', 'hyphen')).toEqual(['hyphen']);
		expect(found('a­', '­')).toEqual([]);
	});

	it('returns no matches for an empty query', () => {
		expect(compileFindPattern('')).toBeNull();
		expect(findMatches('abc', '')).toEqual([]);
	});

	it('does not split a grapheme cluster', () => {
		expect(found('é', 'e')).toEqual([]);
		expect(found('é', 'é')).toEqual(['é']);
	});

	it('reports offsets into the text', () => {
		const [match] = findMatches('xx hello', 'hello');
		expect(match).toMatchObject({ start: 3, end: 8, text: 'hello', groups: ['hello'] });
	});
});

describe('findMatches: Find codes', () => {
	it('matches paragraph marks, tabs, breaks and special characters', () => {
		expect(found('one\rtwo', 'one^ptwo')).toEqual(['one\rtwo']);
		expect(found('a\tb', 'a^tb')).toEqual(['a\tb']);
		expect(found('a\vb', 'a^lb')).toEqual(['a\vb']);
		expect(found('a\fb', 'a^mb')).toEqual(['a\fb']);
		expect(found('a b', 'a^sb')).toEqual(['a b']);
		expect(found('a‑b', 'a^~b')).toEqual(['a‑b']);
		expect(found('a­b', 'a^-b')).toEqual(['a­b']);
		expect(found('a^b', 'a^^b')).toEqual(['a^b']);
	});

	it('matches any character, digit, letter and white space', () => {
		expect(found('a1', '^?^?')).toEqual(['a1']);
		expect(found('a1 b2', '^$^#')).toEqual(['a1', 'b2']);
		expect(found('a b\tc', 'a^wb')).toEqual(['a b']);
		expect(found('a\rb', 'a^?b')).toEqual([]);
	});

	it('matches a surrogate pair as one character', () => {
		expect(found('a😀b', 'a^?b')).toEqual(['a😀b']);
	});

	it('rejects an unknown code', () => {
		expect(() => findMatches('abc', 'a^zb')).toThrow(FindPatternError);
		expect(() => findMatches('abc', 'a^')).toThrow(FindPatternError);
	});
});

describe('findMatches: wildcards', () => {
	it('matches ? as one character and * as any run, shortest first', () => {
		expect(wild('cat cut coat', 'c?t')).toEqual(['cat', 'cut']);
		expect(wild('bad bed', 'b*d')).toEqual(['bad', 'bed']);
	});

	it('does not let ? or * cross a paragraph mark', () => {
		expect(wild('a\rb', 'a*b')).toEqual([]);
		expect(wild('a\rb', 'a?b')).toEqual([]);
	});

	it('treats wildcard-only characters as literals in plain search', () => {
		expect(found('c?t', 'c?t')).toEqual(['c?t']);
		expect(found('cat', 'c?t')).toEqual([]);
	});

	it('matches @ as one or more of the previous atom and {n,m} as a count', () => {
		expect(wild('caaat ct', 'ca@t')).toEqual(['caaat']);
		expect(wild('a aa aaa aaaa', 'a{2,3}')).toEqual(['aa', 'aaa', 'aaa']);
		expect(wild('aa aaaa a', 'a{2}')).toEqual(['aa', 'aa', 'aa']);
		expect(wild('a aa aaa', 'a{2,}')).toEqual(['aa', 'aaa']);
	});

	it('matches word starts and ends with < and >', () => {
		expect(wild('the other bathe', '<the')).toEqual(['the']);
		expect(wild('the other bathe', 'the>')).toEqual(['the', 'the']);
		expect(wild('the other', '<the>')).toEqual(['the']);
	});

	it('matches character sets, ranges and negated sets', () => {
		expect(wild('bat cat hat', '[bc]at')).toEqual(['bat', 'cat']);
		expect(wild('b1 bx', 'b[0-9]')).toEqual(['b1']);
		expect(wild('bat cat hat', '[!ch]at')).toEqual(['bat']);
		expect(wild('a-b', '[-]')).toEqual(['-']);
	});

	it('matches with groups and reports what each group matched', () => {
		const [match] = findMatches('John Smith', '(J*) (S*h)', { wildcards: true });
		expect(match?.groups).toEqual(['John Smith', 'John', 'Smith']);
	});

	it('escapes wildcard characters with a backslash', () => {
		expect(wild('a?b axb', 'a\\?b')).toEqual(['a?b']);
		expect(wild('a*b', 'a\\*b')).toEqual(['a*b']);
		expect(wild('[x]', '\\[x\\]')).toEqual(['[x]']);
	});

	it('understands Find codes inside a wildcard pattern', () => {
		expect(wild('a\rb', 'a^pb')).toEqual(['a\rb']);
		expect(wild('ab12', 'a^#{2}')).toEqual([]);
		expect(wild('a12', 'a^#{2}')).toEqual(['a12']);
	});

	it('disables whole-word, prefix, suffix and ignore options with wildcards', () => {
		expect(findMatches('there', 't?e', { wildcards: true, wholeWords: true })).toHaveLength(1);
		expect(findMatches('ab', 'a?', { wildcards: true, ignoreWhitespace: true })).toHaveLength(1);
	});

	it('keeps Match case', () => {
		expect(findMatches('ABC abc', '[a-z]bc', { wildcards: true, matchCase: true })).toHaveLength(1);
	});

	it('rejects malformed patterns', () => {
		for (const pattern of [
			'(a',
			'a)',
			'[abc',
			'[]',
			'[z-a]',
			'a{2',
			'a{x}',
			'a{3,2}',
			'@',
			'a\\',
			'{2}',
			'(a)(b)(c)(d)(e)(f)(g)(h)(i)(j)',
		]) {
			expect(() => wildcardSource(pattern), pattern).toThrow(FindPatternError);
		}
	});
});

describe('replacementParts', () => {
	const match = (text: string, groups: string[] = [text]) => ({ text, groups });

	it('inserts the found text with ^&', () => {
		expect(replacementParts('[^&]', match('x'))).toEqual([{ kind: 'text', text: '[x]' }]);
		expect(replacementParts('$&', match('x'))).toEqual([{ kind: 'text', text: '$&' }]);
	});

	it('expands Find codes into breaks and special characters', () => {
		expect(replacementParts('a^pb^tc^ld^me^-f^~g^sh^^', match('x'))).toEqual([
			{ kind: 'text', text: 'a' },
			{ kind: 'paragraph' },
			{ kind: 'text', text: 'b' },
			{ kind: 'tab' },
			{ kind: 'text', text: 'c' },
			{ kind: 'lineBreak' },
			{ kind: 'text', text: 'd' },
			{ kind: 'pageBreak' },
			{ kind: 'text', text: 'e' },
			{ kind: 'optionalHyphen' },
			{ kind: 'text', text: 'f' },
			{ kind: 'nonBreakingHyphen' },
			{ kind: 'text', text: 'g' },
			{ kind: 'nonBreakingSpace' },
			{ kind: 'text', text: 'h^' },
		]);
	});

	it('turns a typed line break into a line break', () => {
		expect(replacementParts('a\r\nb\nc', match('x'))).toEqual([
			{ kind: 'text', text: 'a' },
			{ kind: 'lineBreak' },
			{ kind: 'text', text: 'b' },
			{ kind: 'lineBreak' },
			{ kind: 'text', text: 'c' },
		]);
	});

	it('expands \\1 to \\9 to the groups only in wildcard mode', () => {
		const groups = ['John Smith', 'John', 'Smith'];
		expect(replacementParts('\\2, \\1', match('John Smith', groups), true)).toEqual([
			{ kind: 'text', text: 'Smith, John' },
		]);
		expect(replacementParts('\\2', match('John Smith', groups))).toEqual([
			{ kind: 'text', text: '\\2' },
		]);
		expect(() => replacementParts('\\4', match('John Smith', groups), true)).toThrow(
			FindPatternError,
		);
	});

	it('rejects an unknown code', () => {
		expect(() => replacementParts('^z', match('x'))).toThrow(FindPatternError);
	});
});

describe('compileFindPattern', () => {
	it('reports the group count of a wildcard pattern', () => {
		expect(compileFindPattern('(a)(b*)', { wildcards: true })?.groupCount).toBe(2);
		expect(compileFindPattern('ab')?.groupCount).toBe(0);
	});
});
