import { describe, expect, it } from 'vitest';
import { compareText, foldText, textKey } from './text-compare';

const ALPHABET = [
	'a',
	'B',
	'z',
	'0',
	'9',
	' ',
	'-',
	"'",
	'!',
	'~',
	'é',
	'ß',
	'SS',
	'æ',
	'Σ',
	'ς',
	'Ａ',
	'ア',
	'ﬁ',
];

function random(seed: number) {
	let state = seed;
	return () => {
		state = (state * 1103515245 + 12345) & 0x7fffffff;
		return state / 0x80000000;
	};
}

function text(next: () => number): string {
	let out = '';
	const length = Math.floor(next() * 6);
	for (let i = 0; i < length; i++) out += ALPHABET[Math.floor(next() * ALPHABET.length)];
	return out;
}

describe('linguistic text comparison', () => {
	it('is antisymmetric, agrees with textKey on equality and keeps the ASCII fast path exact', () => {
		const next = random(7);
		for (let trial = 0; trial < 5000; trial++) {
			const a = text(next);
			const b = text(next);
			const c = compareText(a, b);
			expect(Math.sign(compareText(b, a)) || 0, `${a} | ${b}`).toBe(-Math.sign(c) || 0);
			expect(c === 0, `${a} | ${b}`).toBe(textKey(a) === textKey(b));
			// A shared accented first letter forces the full comparison from the start.
			expect(Math.sign(compareText(`é${a}`, `é${b}`)), `${a} | ${b}`).toBe(Math.sign(c));
		}
	});

	it('sorts consistently', () => {
		const next = random(11);
		const items = Array.from({ length: 300 }, () => text(next)).sort(compareText);
		for (let i = 1; i < items.length; i++)
			expect(compareText(items[i - 1]!, items[i]!)).toBeLessThanOrEqual(0);
	});

	it('folds character by character without changing the length', () => {
		expect(foldText('STRAẞE ΣΑΣ ς İ')).toBe('straße σασ σ İ');
	});
});
