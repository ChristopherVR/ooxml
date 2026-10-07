import { expect, it } from 'vitest';
import { wrapStyledRuns } from './wrap-styled-runs';

const measure = (text: string) => [...text].length;
const strings = (lines: { text: string }[][]) =>
	lines.map((line) => line.map((run) => run.text).join(''));

it('keeps words crossing style boundaries together and does not mutate input', () => {
	const runs = [
		{ text: 'a Hel', bold: true },
		{ text: 'lo world', bold: false },
	];
	const original = structuredClone(runs);
	const lines = wrapStyledRuns(runs, 5, measure);
	expect(strings(lines)).toEqual(['a', 'Hello', 'world']);
	expect(lines[1]).toEqual([
		{ text: 'Hel', bold: true },
		{ text: 'lo', bold: false },
	]);
	expect(runs).toEqual(original);
});

it('measures combined fragments to retain kerning at word boundaries', () => {
	const lines = wrapStyledRuns([{ text: 'A V' }], 2, (text) => (text === 'A V' ? 2 : text.length));
	expect(strings(lines)).toEqual(['A V']);
});

it('retains explicit blank lines and treats nonbreaking spaces as part of a word', () => {
	expect(strings(wrapStyledRuns([{ text: '\na\r\n\nb\n' }], 10, measure))).toEqual([
		'',
		'a',
		'',
		'b',
		'',
	]);
	expect(strings(wrapStyledRuns([{ text: 'a\u00a0b c' }], 2, measure))).toEqual(['a\u00a0b', 'c']);
});

it('splits oversized words only when requested and keeps graphemes intact', () => {
	const runs = [{ text: 'ab😀e\u0301cd', color: 'red' }];
	expect(strings(wrapStyledRuns(runs, 2, measure))).toEqual(['ab😀e\u0301cd']);
	const lines = wrapStyledRuns(runs, 2, measure, { breakWords: true });
	expect(strings(lines)).toEqual(['ab', '😀', 'e\u0301', 'cd']);
	expect(lines.flat().every((run) => run.color === 'red')).toBe(true);
});

it('allows one oversized grapheme and trims separators at automatic line edges', () => {
	expect(
		strings(wrapStyledRuns([{ text: '  abc   de  ' }], 2, measure, { breakWords: true })),
	).toEqual(['ab', 'c', 'de']);
	expect(strings(wrapStyledRuns([{ text: '😀' }], 0, measure, { breakWords: true }))).toEqual([
		'😀',
	]);
});
