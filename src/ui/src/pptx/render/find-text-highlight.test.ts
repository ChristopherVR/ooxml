import { describe, expect, it } from 'vitest';

import type { FindResult } from './find-replace';
import {
	findMatchMarkStyle,
	findTextHighlightsByElement,
	highlightFindMatches,
	splitTextByFindHighlights,
} from './find-text-highlight';
import type { RenderParagraph } from './paragraph-types';

const match = (over: Partial<FindResult>): FindResult => ({
	slideIndex: 0,
	elementId: 'a',
	segmentIndex: 0,
	startOffset: 0,
	length: 3,
	...over,
});

const paragraph = (runs: RenderParagraph['runs']): RenderParagraph => ({
	runs,
	bulletStyle: {},
});

describe('findTextHighlightsByElement', () => {
	it('groups the slide matches by element and segment and marks the current one', () => {
		const matches = [
			match({}),
			match({ segmentIndex: 1, startOffset: 4 }),
			match({ elementId: 'b' }),
			match({ slideIndex: 1 }),
		];
		const map = findTextHighlightsByElement(matches, 1, 0);
		expect([...map.keys()]).toEqual(['a', 'b']);
		expect(map.get('a')?.get(0)).toEqual([{ startOffset: 0, length: 3, isCurrent: false }]);
		expect(map.get('a')?.get(1)).toEqual([{ startOffset: 4, length: 3, isCurrent: true }]);
	});

	it('leaves out matches outside a shape own text', () => {
		const matches = [
			match({ notes: true, elementId: '' }),
			match({ tableCell: { rowIndex: 0, colIndex: 0 } }),
			match({ smartArtNodeId: 'n' }),
			match({ elementPath: [0] }),
			match({ chartTitle: true }),
			match({ chartSeriesIndex: 0 }),
		];
		expect(findTextHighlightsByElement(matches, 0, 0).size).toBe(0);
	});
});

describe('splitTextByFindHighlights', () => {
	it('splits around the matches', () => {
		expect(
			splitTextByFindHighlights('one two one', [
				{ startOffset: 0, length: 3, isCurrent: false },
				{ startOffset: 8, length: 3, isCurrent: true },
			]),
		).toEqual([
			{ text: 'one', match: 'match' },
			{ text: ' two ' },
			{ text: 'one', match: 'current' },
		]);
	});

	it('does not repeat the characters of overlapping matches and clips to the text', () => {
		const chunks = splitTextByFindHighlights('aaa', [
			{ startOffset: 0, length: 2, isCurrent: false },
			{ startOffset: 1, length: 5, isCurrent: true },
		]);
		expect(chunks.map((chunk) => chunk.text).join('')).toBe('aaa');
		expect(chunks).toEqual([
			{ text: 'aa', match: 'match' },
			{ text: 'a', match: 'current' },
		]);
	});

	it('returns the whole text when nothing matches', () => {
		expect(splitTextByFindHighlights('abc', undefined)).toEqual([{ text: 'abc' }]);
	});
});

describe('highlightFindMatches', () => {
	it('returns the paragraphs unchanged without highlights', () => {
		const paragraphs = [paragraph([{ text: 'x', style: {} }])];
		expect(highlightFindMatches(paragraphs, undefined)).toBe(paragraphs);
	});

	it('splits per-word runs at match boundaries using their segment offsets', () => {
		const paragraphs = [
			paragraph([
				{ text: 'Hello ', style: { color: 'red' }, segmentIndex: 0, charStart: 0 },
				{ text: 'world', style: { color: 'red' }, segmentIndex: 0, charStart: 6 },
			]),
		];
		const map = findTextHighlightsByElement([match({ startOffset: 4, length: 4 })], 0, 0).get('a');
		const runs = highlightFindMatches(paragraphs, map)[0]?.runs ?? [];
		expect(runs.map((run) => [run.text, run.findMatch, run.charStart])).toEqual([
			['Hell', undefined, 0],
			['o ', 'current', 4],
			['wo', 'current', 6],
			['rld', undefined, 8],
		]);
		expect(runs[1]?.style).toEqual({ color: 'red' });
	});

	it('slices script-font pieces with the run', () => {
		const paragraphs = [
			paragraph([
				{
					text: 'ab日本',
					style: {},
					segmentIndex: 0,
					charStart: 0,
					scriptRuns: [{ text: 'ab' }, { text: '日本', style: { fontFamily: 'X' } }],
				},
			]),
		];
		const map = findTextHighlightsByElement([match({ startOffset: 1, length: 2 })], -1, 0).get('a');
		const runs = highlightFindMatches(paragraphs, map)[0]?.runs ?? [];
		expect(runs[1]?.findMatch).toBe('match');
		expect(runs[1]?.scriptRuns).toEqual([
			{ text: 'b' },
			{ text: '日', style: { fontFamily: 'X' } },
		]);
	});

	it('marks a ruby or tab run whole instead of splitting it', () => {
		const paragraphs = [
			paragraph([
				{
					text: 'kanji',
					style: {},
					segmentIndex: 0,
					charStart: 0,
					ruby: { text: 'k', style: {} },
				},
			]),
		];
		const map = findTextHighlightsByElement([match({ startOffset: 1, length: 1 })], 0, 0).get('a');
		const runs = highlightFindMatches(paragraphs, map)[0]?.runs ?? [];
		expect(runs).toHaveLength(1);
		expect(runs[0]?.findMatch).toBe('current');
	});
});

describe('findMatchMarkStyle', () => {
	it('tells the current match apart', () => {
		expect(findMatchMarkStyle('current').backgroundColor).not.toBe(
			findMatchMarkStyle('match').backgroundColor,
		);
	});
});
