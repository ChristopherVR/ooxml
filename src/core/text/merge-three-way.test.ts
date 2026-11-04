import { describe, expect, it } from 'vitest';
import { mergeTextThreeWay } from './merge-three-way';

describe('three-way text merge', () => {
	it.each([
		['Hello world', 'Hello brave world', 'Hello world!', 'Hello brave world!'],
		['abc', 'aBc', 'aXbc', 'aXBc'],
		['abc', 'ac', 'abc!', 'ac!'],
		['one two three', 'ONE two three', 'one two THREE', 'ONE two THREE'],
		['a😀z', 'a😀!z', 'A😀z', 'A😀!z'],
		['ab', 'aXb', 'aYb', 'aXYb'],
		['ab', 'aXb', 'aXb', 'aXb'],
	])('retains independent edits to %s', (base, local, remote, result) => {
		expect(mergeTextThreeWay(base, local, remote)).toEqual({ status: 'merged', text: result });
		expect(mergeTextThreeWay(base, remote, local)).toEqual({ status: 'merged', text: result });
	});
	it('declines incompatible replacements of the same original text', () => {
		expect(mergeTextThreeWay('cat', 'dog', 'fox')).toEqual({ status: 'unresolved' });
	});
});
