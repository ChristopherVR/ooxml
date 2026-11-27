import { describe, expect, it } from 'vitest';
import { parseTransform, transformCalls } from './pdf-svg-style';

/** The pattern the scanner replaces, kept here to pin its behaviour. */
const regexCalls = (value: string): [string, string][] =>
	[...value.matchAll(/(\w+)\s*\(([^)]*)\)/g)].map((match) => [match[1]!, match[2]!]);

describe('transformCalls', () => {
	it.each([
		'',
		'translate(10 20)',
		'translate(10,20) scale(2)',
		'matrix(1 0 0 1 5 5)rotate(45 1 1)',
		'  scale\t( 2 )  ',
		'rotate (30)skewX(10) skewY(5)',
		'a(b(c)d)',
		'x(1',
		'x 1) y(2)',
		'-translate(1)',
		'(1) scale(2)',
		'translate(1\n2)',
		'_t9(1)',
		'é(1) ab(2)',
		'000 0((0(()',
		'scale(2) trailing',
	])('finds the calls the regular expression finds in %j', (value) => {
		expect(transformCalls(value)).toEqual(regexCalls(value));
	});

	it('matches the regular expression on generated transform-like text', () => {
		const alphabet = ['a', '0', '_', ' ', '\t', '(', ')', ',', '-', '.'];
		let seed = 7;
		const next = () => (seed = (seed * 1103515245 + 12345) % 2147483648);
		for (let sample = 0; sample < 2000; sample++) {
			let value = '';
			const length = next() % 24;
			for (let index = 0; index < length; index++) value += alphabet[next() % alphabet.length];
			expect(transformCalls(value), value).toEqual(regexCalls(value));
		}
	});

	it('reads a long word run without a call in linear time', () => {
		const started = performance.now();
		expect(transformCalls('0'.repeat(200_000))).toEqual([]);
		expect(transformCalls(`0(${'0(('.repeat(50_000)}`)).toEqual([]);
		expect(performance.now() - started).toBeLessThan(500);
	});

	it('keeps parseTransform results', () => {
		expect(parseTransform('translate(3 4) scale(2)')).toEqual([
			[1, 0, 0, 1, 3, 4],
			[2, 0, 0, 2, 0, 0],
		]);
		expect(parseTransform(null)).toEqual([]);
	});
});
